# core/scan_main.py
"""
메인 반이동 파이프라인의 스캔 전용 모듈.

책임 범위:
  - 입력 파일 탐색
  - 학교 정보 확인
  - 학생명부 / 신입생 / 전입생 / 전출생 / 교직원 파일 존재 확인
  - 헤더 행 / 데이터 시작 행 / 주요 컬럼 자동 감지
  - 명부 기반 학년/반 구조 분석
  - scan_pipeline() 실행 결과 반환

이 모듈은 "실행(run)"이 아니라 "사전 분석(scan)"만 담당.

공개 API:
  ScanResult
  scan_pipeline(work_root, school_name, school_start_date, work_date, roster_basis_date) -> ScanResult
  get_project_dirs(work_root) -> Dict[str, Path]
  get_school_domain(db_dir, school_name) -> Optional[str]
  detect_input_layout(xlsx_path, kind) -> Dict[str, Any]
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Sequence
from collections import Counter, defaultdict

from core.utils import normalize_text, text_contains
from core.events import (
    CoreEvent, RowMark,
    roster_xls_format, school_not_in_roster,
    school_folder_not_found, school_folder_ambiguous, no_input_files, input_xls_format,
    missing_header, missing_data_start, empty_data, missing_required_col,
    grade_format_warn, class_format_warn, name_format_warn, empty_required_field, empty_row, merged_cell,
    multiple_sheets, school_kind_unknown, roster_not_found, roster_date_mismatch,
    open_date_missing, duplicate_input_file,
)

from core.xlsx_db import (
    search_schools_in_xlsx as _xlsx_search_schools,
    load_school_names_from_xlsx as _xlsx_load_names,
    get_school_domain_from_xlsx as _xlsx_get_domain,
    school_exists_in_xlsx as _xlsx_school_exists,
)

from core.common import (
    get_project_dirs,
    ensure_xlsx_only,
    safe_load_workbook,
    get_first_sheet_with_warning,
    warn_if_multi_sheet,
    header_map,
    _build_header_slot_map,
    _detect_header_row_generic,
    normalize_name,
    normalize_name_key,
    load_roster_sheet,
    parse_class_str,
    extract_id_prefix4,
    school_kind_from_name,
    school_profile_from_name,
    FRESHMEN_HEADER_SLOTS,
    TRANSFER_HEADER_SLOTS,
    WITHDRAW_HEADER_SLOTS,
    TEACHER_HEADER_SLOTS,
    RosterInfo,
)

# =========================
# Result types
# =========================
@dataclass
class ScanResult:
    # 기본 상태
    ok: bool = False
    logs: List[str] = field(default_factory=list)

    # 학교/연도 정보
    school_name: str = ""
    year_str: str = ""
    year_int: int = 0

    # 경로들
    project_root: Path = Path(".")
    input_dir: Path = Path(".")
    output_dir: Path = Path(".")
    template_register: Optional[Path] = None
    template_notice: Optional[Path] = None
    roster_xlsx_path: Optional[Path] = None

    # 인풋 파일
    freshmen_file: Optional[Path] = None
    teacher_file: Optional[Path] = None
    transfer_file: Optional[Path] = None
    withdraw_file: Optional[Path] = None

    # 학생명부 관련
    need_roster: bool = False              # # 전입/전출이 있거나, 신입생 파일에 1학년 외 학년이 있으면 True
    roster_path: Optional[Path] = None
    roster_year: Optional[int] = None
    roster_info: Optional[RosterInfo] = None
    roster_basis_date: Optional[date] = None  # 학생명부 기준일(파일 수정일 or 사용자가 수정한 값)

    # UI 플래그
    needs_open_date: bool = False          # 전출 있으면 True → 개학일 필요
    roster_date_mismatch: bool = False     # 작업일 ≠ 명부 기준일이면 True → UI에서 수정 옵션 제공
    missing_fields: List[str] = field(default_factory=list)
    can_execute: bool = False
    can_execute_after_input: bool = False
    school_profile_mode: str = "single"
    school_kind_needs_choice: bool = False
    grade_rule_max_grade: int = 6

    # UI용 스캔 메타 (파일별 미리보기 dict)
    freshmen: Optional[Dict[str, Any]] = None
    transfer_in: Optional[Dict[str, Any]] = None
    transfer_out: Optional[Dict[str, Any]] = None
    teachers: Optional[Dict[str, Any]] = None
    roster: Optional[Dict[str, Any]] = None

    # 구조화된 판정 결과 — bridge/UI는 이것만 참조
    events:    List[Any] = field(default_factory=list)  # List[CoreEvent]
    row_marks: List[Any] = field(default_factory=list)  # List[RowMark]


# =========================
# Constants / keywords
# =========================
FRESHMEN_KEYWORDS = ["신입생", "신입", "1학년"]
TEACHER_KEYWORDS  = ["교사", "교원", "교직원"]
TRANSFER_KEYWORDS = ["전입생", "전입"]
WITHDRAW_KEYWORDS = ["전출생", "전출"]


# =========================
# L0. Path / file utils (main_scan 전용)
# =========================


def find_single_input_file(input_dir: Path, keywords: Sequence[str], *, file_key: str | None = None, kind_label: str | None = None) -> Optional[Path]:
    if not input_dir.exists():
        return None

    kw_list: List[str] = []
    for k in keywords:
        k = "" if k is None else str(k).strip()
        if k:
            kw_list.append(k)

    if not kw_list:
        return None

    candidates: List[Path] = []
    xls_candidates: List[Path] = []  # .xls(구형) 파일 감지용

    for p in input_dir.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if not any(text_contains(p.name, kw) for kw in kw_list):
            continue
        if p.suffix.lower() == ".xlsx":
            candidates.append(p)
        elif p.suffix.lower() == ".xls":
            xls_candidates.append(p)

    # .xlsx 없고 .xls만 있으면 명시적 에러
    if len(candidates) == 0 and xls_candidates:
        names = ", ".join(p.name for p in xls_candidates)
        raise ValueError(
            f"__XLS_FORMAT__ .xls 파일이 감지되었습니다: {names}\n"
            ".xls 형식은 지원하지 않습니다. Excel에서 .xlsx로 저장 후 다시 시도해 주세요."
        )

    if len(candidates) == 0:
        return None
    if len(candidates) > 1:
        raise ValueError(f"__DUPLICATE__ 입력 파일이 2개 이상 감지되었습니다: {[c.name for c in candidates]}")
    return candidates[0]


# =========================
# L0.5 DB helpers
# =========================


def load_all_school_names(
    roster_xlsx: Optional[Path],
    col_map: Optional[Dict[str, Any]] = None,
) -> List[str]:
    """명단 xlsx에서 전체 학교명 목록 로드."""
    if not roster_xlsx:
        return []
    names = _xlsx_load_names(Path(roster_xlsx), col_map=col_map)
    cleaned = []
    seen: set = set()
    for name in names or []:
        s = str(name).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        cleaned.append(s)
    return cleaned


def get_school_domain(
    roster_xlsx: Optional[Path],
    school_name: str,
    col_map: Optional[Dict[str, Any]] = None,
) -> Optional[str]:
    """명단 xlsx에서 학교 도메인 조회."""
    if not roster_xlsx:
        return None
    return _xlsx_get_domain(Path(roster_xlsx), school_name, col_map=col_map)


# =========================
# L1. Input scan / layout detection helpers
# =========================
def find_templates(format_dir: Path) -> Tuple[Optional[Path], Optional[Path], List[str]]:
    """
    [templates] 폴더 템플릿 2개 식별:
    - 등록 템플릿: 파일명에 '등록' 포함
    - 안내 템플릿: 파일명에 '안내' 포함
    """
    format_dir = Path(format_dir).resolve()
    if not format_dir.exists():
        return None, None, [f"[ERROR] templates 폴더를 찾을 수 없습니다."]

    xlsx_files = [
        p for p in format_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
    ]
    if not xlsx_files:
        return None, None, [f"[ERROR] templates 폴더에 xlsx 파일이 없습니다."]

    import unicodedata as _ud
    reg = [p for p in xlsx_files if "등록" in _ud.normalize('NFC', p.stem)]
    notice = [p for p in xlsx_files if "안내" in _ud.normalize('NFC', p.stem)]

    errors: List[str] = []
    if len(reg) == 0:
        errors.append("[ERROR] templates 폴더에서 '등록' 템플릿을 찾을 수 없습니다. (파일명에 '등록' 포함)")
    elif len(reg) > 1:
        errors.append("[ERROR] templates 폴더에 '등록' 템플릿이 여러 개 있습니다.")

    if len(notice) == 0:
        errors.append("[ERROR] templates 폴더에서 '안내' 템플릿을 찾을 수 없습니다. (파일명에 '안내' 포함)")
    elif len(notice) > 1:
        errors.append("[ERROR] templates 폴더에 '안내' 템플릿이 여러 개 있습니다.")

    if errors:
        return None, None, errors

    return reg[0], notice[0], []


def choose_template_register(format_dir: Path, year_str: str = "") -> Path:
    reg, notice, errors = find_templates(format_dir)
    if errors:
        raise ValueError(errors[0])
    assert reg is not None
    return reg


def choose_template_notice(format_dir: Path, year_str: str = "") -> Path:
    reg, notice, errors = find_templates(format_dir)
    if errors:
        raise ValueError(errors[-1])
    assert notice is not None
    return notice


# --- header cell normalize / detection ---


# =========================
# 헤더 행 자동 감지
# =========================
# HEADER_SLOTS 상수는 common.py에 정의되어 있음.
# scan(헤더 감지)과 run(데이터 읽기) 양쪽에서 공유하는 상수이므로
# 어느 한쪽에 두지 않고 common에 정의한다.

def detect_header_row_freshmen(ws) -> int:
    """신입생 파일에서 헤더 행 번호를 자동 감지한다."""
    return _detect_header_row_generic(ws, FRESHMEN_HEADER_SLOTS,
                                      max_search_row=15, max_col=10, min_match_slots=3)


def detect_header_row_transfer(ws) -> int:
    """전입생 파일에서 헤더 행 번호를 자동 감지한다."""
    return _detect_header_row_generic(ws, TRANSFER_HEADER_SLOTS,
                                      max_search_row=15, max_col=10, min_match_slots=3)


def detect_header_row_withdraw(ws) -> int:
    """전출생 파일에서 헤더 행 번호를 자동 감지한다."""
    return _detect_header_row_generic(ws, WITHDRAW_HEADER_SLOTS,
                                      max_search_row=15, max_col=10, min_match_slots=3)


def detect_header_row_teacher(ws) -> int:
    """교사 파일에서 헤더 행 번호를 자동 감지한다."""
    return _detect_header_row_generic(ws, TEACHER_HEADER_SLOTS,
                                      max_search_row=15, max_col=10, min_match_slots=3)



# =========================
# example row detection (예시 + 데이터 시작 행)
# =========================
EXAMPLE_NAMES_RAW = ["홍길동", "이순신", "유관순", "임꺽정"]
EXAMPLE_NAMES_NORM = {normalize_text(n) for n in EXAMPLE_NAMES_RAW}
EXAMPLE_KEYWORDS = ["예시"]


def _row_is_empty(ws, row: int, max_col: Optional[int] = None) -> bool:
    if max_col is None:
        max_col = ws.max_column or 1
    for c in range(1, max_col + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and str(v).strip() != "":
            return False
    return True


def _row_has_example_keyword(ws, row: int, max_col: Optional[int] = None) -> bool:
    if max_col is None:
        max_col = ws.max_column or 1
    for c in range(1, max_col + 1):
        v = ws.cell(row=row, column=c).value
        if v is None:
            continue
        s = normalize_text(str(v))
        if not s:
            continue
        for kw in EXAMPLE_KEYWORDS:
            if kw in s:
                return True
    return False


def _cell_is_example_name(value: Any) -> bool:
    if value is None:
        return False
    s = normalize_text(str(value))
    return bool(s) and s in EXAMPLE_NAMES_NORM


def detect_example_and_data_start(
    ws,
    header_row: int,
    name_col: int,
    max_search_row: Optional[int] = None,
    max_col: Optional[int] = None,
) -> Tuple[List[int], int]:
    """
    헤더 아래에서 예시 행(0개 이상)과 실제 데이터 시작 행을 자동 감지한다.
    """
    if max_search_row is None:
        max_search_row = ws.max_row

    example_rows: List[int] = []
    r = header_row + 1

    while r <= max_search_row:
        if _row_is_empty(ws, r, max_col=max_col):
            r += 1
            continue

        if _row_has_example_keyword(ws, r, max_col=max_col):
            example_rows.append(r)
            r += 1
            continue

        v_name = ws.cell(row=r, column=name_col).value
        if _cell_is_example_name(v_name):
            example_rows.append(r)
            r += 1
            continue

        return example_rows, r

    raise ValueError(
        f"[ERROR] 데이터 시작 행을 찾을 수 없습니다. 헤더 행 아래에 실제 데이터가 있는지 확인해 주세요."
    )


def detect_input_layout(xlsx_path: Path, kind: str) -> Dict[str, Any]:
    """
    UI에서 인풋 파일 구조를 미리 보여줄 때 사용.
    kind: 'freshmen' | 'transfer' | 'withdraw' | 'teacher'
    반환:
      {
        "header_row": int,
        "example_rows": [int, ...],
        "data_start_row": int,
        "_wb": Workbook,   # 내부 재사용용 — 호출자가 꺼내 쓰고 닫는다
        "_ws": Worksheet,
      }
    read_only=True로 열어 속도를 최적화한다.
    (병합셀 정보는 read_only 모드에서 불가 — validate에서 스킵됨)
    """
    ensure_xlsx_only(xlsx_path)
    wb = safe_load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = get_first_sheet_with_warning(wb, xlsx_path.name)

    kind_norm = (kind or "").strip().lower()

    if kind_norm == "freshmen":
        header_row = detect_header_row_freshmen(ws)
        slot_cols = _build_header_slot_map(ws, header_row, FRESHMEN_HEADER_SLOTS)
        name_col = slot_cols.get("name", 5)

    elif kind_norm == "transfer":
        header_row = detect_header_row_transfer(ws)
        slot_cols = _build_header_slot_map(ws, header_row, TRANSFER_HEADER_SLOTS)
        name_col = slot_cols.get("name", 5)

    elif kind_norm == "withdraw":
        header_row = detect_header_row_withdraw(ws)
        slot_cols = _build_header_slot_map(ws, header_row, WITHDRAW_HEADER_SLOTS)
        name_col = slot_cols.get("name", 4)

    elif kind_norm == "teacher":
        header_row = detect_header_row_teacher(ws)
        slot_cols = _build_header_slot_map(ws, header_row, TEACHER_HEADER_SLOTS)
        name_col = slot_cols.get("name", 3)

    else:
        raise ValueError(f"[ERROR] 지원하지 않는 파일 종류입니다: {kind}")

    example_rows, data_start_row = detect_example_and_data_start(
        ws,
        header_row=header_row,
        name_col=name_col,
    )

    return {
        "header_row": header_row,
        "example_rows": example_rows,
        "data_start_row": data_start_row,
        "_wb": wb,
        "_ws": ws,
    }

def _sheet_headers(ws, header_row: int, max_col: int = 12) -> List[str]:
    headers = []
    limit = min(ws.max_column or 1, max_col)
    for c in range(1, limit + 1):
        v = ws.cell(header_row, c).value
        headers.append("" if v is None else str(v).strip())
    return headers


def _sheet_preview_rows(
    ws,
    data_start_row: int,
    required_cols: Optional[Dict[str, int]] = None,
    max_col: int = 12,
    limit: int = 300,
) -> List[List[str]]:
    """
    미리보기용 행 추출.

    원칙:
    - 종료 판단은 No가 아니라 required_cols(필수 열) 기준으로 한다.
    - 필수 열이 10행 이상 연속 공백이면 데이터 끝으로 간주한다.
    - 데이터 끝(last_data_row) 이전 행은 빈 행이라도 자리 그대로 유지한다.
      → 중간 빈 행 / No만 있는 행도 표에서 확인 가능해야 함.
    """
    rows: List[List[str]] = []
    col_limit = min(ws.max_column or 1, max_col)

    req_cols = dict(required_cols or {})
    max_req_col = max(req_cols.values()) if req_cols else col_limit
    scan_max_col = max(col_limit, max_req_col)
    MAX_BLANK_STREAK = 10

    last_data_row = data_start_row - 1
    blank_streak = 0
    collected: List[tuple[int, List[str], bool]] = []

    for r in range(data_start_row, ws.max_row + 1):
        all_vals: List[str] = []
        for c in range(1, scan_max_col + 1):
            v = ws.cell(r, c).value
            all_vals.append("" if v is None else str(v).strip())

        if req_cols:
            req_vals = [all_vals[c - 1] if c - 1 < len(all_vals) else "" for c in req_cols.values()]
            has_required = any(v for v in req_vals)
        else:
            has_required = any(v for v in all_vals[:col_limit])

        if has_required:
            last_data_row = r
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak >= MAX_BLANK_STREAK:
                break

        collected.append((r, all_vals[:col_limit], has_required))

    if last_data_row < data_start_row:
        return []

    def _is_blank_or_seq_only(row_vals: List[str]) -> bool:
        nonempty = [i for i, v in enumerate(row_vals) if str(v).strip()]
        return (not nonempty) or nonempty == [0]

    for r, row_vals, has_required in collected:
        if r > last_data_row:
            break
        # 미리보기에서는 필수열이 비어 있고 No(첫 열)만 있는 행, 완전 빈 행은 숨긴다.
        # 종료 판단은 기존처럼 required_cols 기준 10행 연속 공백을 유지한다.
        if (not has_required) and _is_blank_or_seq_only(row_vals):
            continue
        rows.append(row_vals)
        if len(rows) >= limit:
            break

    return rows
    

def _collect_merged_ranges_in_data_area(ws, data_start_row: int) -> List[str]:
    """
    자동감지된 data_start_row 이후 데이터 영역과 겹치는 병합셀만 수집.
    제목/상단 안내 영역 병합셀은 무시.
    ReadOnlyWorksheet면 빈 리스트 반환.
    """
    if not hasattr(ws, "merged_cells"):
        return []

    merged = getattr(ws, "merged_cells", None)
    if not merged:
        return []

    hit_ranges: List[str] = []
    for rng in merged.ranges:
        # 병합범위가 데이터 시작행 아래와 겹치면 경고 대상
        if rng.max_row >= data_start_row:
            hit_ranges.append(str(rng))
    return hit_ranges


def validate_input_sheet_structure(
    ws,
    kind: str,
    header_row: int,
    data_start_row: int,
    required_cols: Dict[str, int],
    allow_blank_class_for_kindergarten: bool = False,
    file_key: str = "global",
) -> Tuple[List[str], List[int], List[CoreEvent], List[RowMark]]:
    """
    입력 시트 구조 검증:
    - 데이터 영역 병합 셀 있으면 경고
    - 데이터 중간 완전 빈 행 있으면 경고
    - 필수 컬럼 중간 빈칸 있으면 경고
    반환: (issues, issue_row_nums, events, row_marks)
    """
    issues: List[str] = []
    issue_row_nums: List[int] = []
    evts: List[CoreEvent] = []
    marks: List[RowMark] = []

    merged_ranges = _collect_merged_ranges_in_data_area(ws, data_start_row)
    if merged_ranges:
        issues.append(
            f"[WARN] {kind} 파일 데이터 영역에 병합된 셀이 있습니다: "
            f"{', '.join(merged_ranges[:10])}"
            + (" ..." if len(merged_ranges) > 10 else "")
        )
        evts.append(merged_cell(file_key, kind))

    # iter_rows로 한 번에 읽어 max_row 오염 및 랜덤접근 느림 방지
    # 연속 빈 행 MAX_BLANK_STREAK개 이상이면 데이터 끝으로 간주하고 조기 종료
    MAX_BLANK_STREAK = 10
    max_col = max(required_cols.values()) if required_cols else 1

    last_data_row  = data_start_row - 1
    blank_streak   = 0
    collected_rows: list = []  # [(row_num, {name: value})]

    for i, row_tuple in enumerate(ws.iter_rows(min_row=data_start_row, max_col=max_col, values_only=True)):
        r = data_start_row + i
        row_vals = {
            name: (row_tuple[c - 1] if c - 1 < len(row_tuple) else None)
            for name, c in required_cols.items()
        }
        collected_rows.append((r, row_vals))

        has_any = any(v is not None and str(v).strip() != "" for v in row_vals.values())
        if has_any:
            last_data_row = r
            blank_streak  = 0
        else:
            blank_streak += 1
            if blank_streak >= MAX_BLANK_STREAK:
                break

    if last_data_row < data_start_row:
        issues.append(f"[ERROR] {kind} 파일에서 데이터 행을 찾을 수 없습니다.")
        evts.append(empty_data(file_key, kind))
        return issues, [], evts, marks

    for r, row_vals in collected_rows:
        if r > last_data_row:
            break
        vals = list(row_vals.values())
        if all(v is None or str(v).strip() == "" for v in vals):
            issues.append(
                f"[WARN] {kind} 파일 {r}행이 비어 있습니다. "
                "중간 빈 행이 있으면 데이터가 잘릴 수 있습니다."
            )
            issue_row_nums.append(r)
            _e, _m = empty_row(file_key, kind, r)
            evts.append(_e); marks.append(_m)

    for r, row_values in collected_rows:
        if r > last_data_row:
            break

        # 완전 빈 행은 중간 빈 행 WARN에서 이미 처리 — 개별 컬럼 체크 스킵
        if all(v is None or str(v).strip() == "" for v in row_values.values()):
            continue

        if allow_blank_class_for_kindergarten:
            grade_v = row_values.get("grade")
            grade_s = "" if grade_v is None else str(grade_v).strip()
            cls_v = row_values.get("class")
            cls_s = "" if cls_v is None else str(cls_v).strip()
            cls_norm = re.sub(r"\s+", "", cls_s)
            grade_norm = re.sub(r"\s+", "", grade_s)
            is_kindergarten = (
                grade_norm in {"유치원", "유치원반", "5세", "6세", "7세", "5세반", "6세반", "7세반"}
                or (not grade_s and cls_norm in {"유치원", "유치원반"})
                or cls_norm in {"유치원", "유치원반"}
            )
            if is_kindergarten and not any(e.code == "KINDERGARTEN_IN_FILE" for e in evts):
                from core.events import kindergarten_in_file as _kif
                evts.append(_kif(file_key, kind))

            for key, v in row_values.items():
                if is_kindergarten and key in ("grade", "class"):
                    continue
                if v is None or str(v).strip() == "":
                    issues.append(f"[WARN] {kind} 파일 {r}행 '{key}' 값이 비어 있습니다.")
                    issue_row_nums.append(r)
                    _e, _m = empty_required_field(file_key, kind, r, key)
                    evts.append(_e); marks.append(_m)
        else:
            for key, v in row_values.items():
                if v is None or str(v).strip() == "":
                    issues.append(f"[WARN] {kind} 파일 {r}행 '{key}' 값이 비어 있습니다.")
                    issue_row_nums.append(r)
                    _e, _m = empty_required_field(file_key, kind, r, key)
                    evts.append(_e); marks.append(_m)

        # 학년/반 열 형식 검증
        KINDERGARTEN = {"유치원", "유치원반", "5세", "6세", "7세", "5세반", "6세반", "7세반"}

        grade_v = row_values.get("grade")
        if grade_v is not None and str(grade_v).strip():
            gs = str(grade_v).strip()
            gs_norm = re.sub(r"\s+", "", gs)
            if gs_norm not in KINDERGARTEN:
                warn_grade = None
                if "-" in gs:
                    warn_grade = f"학년 열에 하이픈(-) 포함 — 학년+반이 합쳐진 것 같습니다: '{gs}'"
                elif gs_norm.endswith("학년"):
                    warn_grade = f"학년 열에 '학년' 글자 포함 — 숫자만 입력해야 합니다: '{gs}'"
                else:
                    try:
                        float(gs)
                    except ValueError:
                        warn_grade = f"학년 열에 숫자가 아닌 값이 있습니다: '{gs}'"
                if warn_grade:
                    issues.append(f"[WARN] {kind} 파일 {r}행 '학년' 열 — {warn_grade}")
                    issue_row_nums.append(r)
                    _e, _m = grade_format_warn(file_key, kind, r, warn_grade)
                    evts.append(_e); marks.append(_m)

        class_v = row_values.get("class")
        if class_v is not None and str(class_v).strip():
            cs = str(class_v).strip()
            cs_norm = re.sub(r"\s+", "", cs)
            warn_class = None
            if "-" in cs:
                warn_class = f"반 열에 하이픈(-) 포함 — 학년+반이 합쳐진 것 같습니다: '{cs}'"
            elif cs_norm.endswith("반"):
                warn_class = f"반 열 끝에 '반' 포함 — 등록 파일에 중복됩니다: '{cs}'"
            if warn_class:
                issues.append(f"[WARN] {kind} 파일 {r}행 '반' 열 — {warn_class}")
                issue_row_nums.append(r)
                _e, _m = class_format_warn(file_key, kind, r, warn_class)
                evts.append(_e); marks.append(_m)

        name_v = row_values.get("name")
        if name_v is not None and str(name_v).strip():
            ns = str(name_v).strip()
            warn_name = None
            if re.fullmatch(r"[0-9]+", ns):
                warn_name = f"이름 열에 숫자만 있습니다: '{ns}'"
            elif len(re.sub(r"\s+", "", ns)) <= 1:
                warn_name = f"이름이 너무 짧습니다(1자 이하): '{ns}'"
            elif re.search(r"[0-9]{3,}", ns):
                warn_name = f"이름 열에 숫자가 3자리 이상 포함되어 있습니다: '{ns}'"
            elif re.search(r"[^\uAC00-\uD7A3\u1100-\u11FF\u3130-\u318Fa-zA-Z\s]", ns):
                warn_name = f"이름에 한글/영문 외 문자가 있습니다: '{ns}'"
            if warn_name:
                issues.append(f"[WARN] {kind} 파일 {r}행 '이름' 열 — {warn_name}")
                issue_row_nums.append(r)
                _e, _m = name_format_warn(file_key, kind, r, warn_name)
                evts.append(_e); marks.append(_m)

    return issues, sorted(set(issue_row_nums)), evts, marks





# =========================
# Roster helpers
# =========================


def analyze_roster_once(roster_ws, input_year: int) -> Dict:
    """
    학생명부 시트를 한 번 순회하여 RosterInfo 구성에 필요한 정보를 추출한다.
    결과는 dict로 반환하며, scan_pipeline 내부에서 RosterInfo 필드로 설정된다.

    [ID prefix 최빈값 분석]
    각 학년별로 학생 아이디 앞 4자리(입학년도)의 최빈값을 구한다.
    이것이 prefix_mode_by_roster_grade가 된다.

    예: 3학년 학생 30명의 아이디 앞 4자리가 대부분 2022이면
        prefix_mode[3] = 2022
        → 3학년 전입생 아이디: "2022{이름}"

    최빈값을 쓰는 이유: 담당자가 특정 전입생에게 잘못된 연도를
    입력했더라도 다수결로 올바른 연도를 추정할 수 있기 때문.

    [roster_time 추정]
    roster_time/ref_grade_shift는 scan_pipeline에서
    명부 기준일과 개학일을 비교하여 확정한다.
    여기서는 "unknown"으로 초기화만 한다.
    """
    hm = header_map(roster_ws, 1)

    need = ["현재반", "이전반", "학생이름", "아이디"]
    for k in need:
        if k not in hm:
            raise ValueError(f"[ERROR] 학생명부에 '{k}' 열이 없습니다.")

    c_class = hm["현재반"]
    c_name  = hm["학생이름"]
    c_id    = hm["아이디"]

    prefixes_by_grade = defaultdict(list)
    name_counter_by_grade = defaultdict(Counter)

    # iter_rows + 연속 빈 행 조기종료 (max_row 오염 대응)
    MAX_BLANK = 20
    blank_streak = 0
    max_col_r = max(c_class, c_name, c_id)
    for row_tuple in roster_ws.iter_rows(min_row=2, max_col=max_col_r, values_only=True):
        clv = row_tuple[c_class - 1] if c_class - 1 < len(row_tuple) else None
        nmv = row_tuple[c_name  - 1] if c_name  - 1 < len(row_tuple) else None
        idv = row_tuple[c_id    - 1] if c_id    - 1 < len(row_tuple) else None
        if clv is None and nmv is None:
            blank_streak += 1
            if blank_streak >= MAX_BLANK:
                break
            continue
        blank_streak = 0
        if clv is None or nmv is None:
            continue

        parsed = parse_class_str(clv)
        if parsed is None:
            continue
        g, _cls = parsed

        nm = normalize_name(nmv)
        if not nm:
            continue
        name_counter_by_grade[g][nm] += 1

        p4 = extract_id_prefix4(idv)
        if p4 is not None:
            prefixes_by_grade[g].append(p4)

    prefix_mode_by_grade = {}
    for g, arr in prefixes_by_grade.items():
        if arr:
            prefix_mode_by_grade[g] = Counter(arr).most_common(1)[0][0]

    # scan 단계에서는 명부에 실제로 존재하는 학년만 기록한다.
    # 신입생 학년 보정은 run_main.build_freshmen_prefix_map()에서
    # 명부 anchor 학년을 기준으로 역산한다.

    # roster_names_by_grade: {학년(int): [이름(str), ...]}
    roster_names_by_grade = {
        g: list(counter.keys())
        for g, counter in name_counter_by_grade.items()
    }

    return RosterInfo(
        roster_time="unknown",   # scan_pipeline에서 날짜 기준으로 확정
        ref_grade_shift=0,
        prefix_mode_by_roster_grade=prefix_mode_by_grade,
        name_count_by_roster_grade=name_counter_by_grade,
        roster_names_by_grade=roster_names_by_grade,
    )

def freshmen_need_roster(
    xlsx_path: Optional[Path],
    input_year: int,
    school_name: str = "",
) -> bool:
    """
    신입생 파일 안에 일반 학년 2~6학년이 하나라도 있으면
    학생명부가 필요하다고 판단한다.
    run_main import 없이 직접 학년 컬럼만 읽어서 판단.
    """
    return bool(freshmen_extra_grades(xlsx_path, input_year, school_name))


def freshmen_extra_grades(
    xlsx_path: Optional[Path],
    input_year: int,
    school_name: str = "",
) -> list:
    """신입생 파일에서 1학년 외 학년 목록을 반환. 없으면 []."""
    if not xlsx_path:
        return []

    wb = None
    try:
        wb = safe_load_workbook(xlsx_path, data_only=True, read_only=True)
        ws = get_first_sheet_with_warning(wb, xlsx_path.name)

        header_row = detect_header_row_freshmen(ws)
        slot_cols = _build_header_slot_map(ws, header_row, FRESHMEN_HEADER_SLOTS)
        grade_col = slot_cols.get("grade")
        name_col = slot_cols.get("name", 5)

        if grade_col is None:
            return []

        _, data_start_row = detect_example_and_data_start(
            ws,
            header_row=header_row,
            name_col=name_col,
        )

        KINDERGARTEN = {"유치원", "유치원반", "5세", "6세", "7세", "5세반", "6세반", "7세반"}
        extra = set()

        # iter_rows + 연속 빈 행 조기종료 (max_row 오염 대응)
        MAX_BLANK = 10
        blank_streak = 0
        for row_tuple in ws.iter_rows(
            min_row=data_start_row,
            min_col=grade_col,
            max_col=grade_col,
            values_only=True,
        ):
            grade_v = row_tuple[0]
            if grade_v is None or str(grade_v).strip() == "":
                blank_streak += 1
                if blank_streak >= MAX_BLANK:
                    break
                continue
            blank_streak = 0

            grade_s = re.sub(r"\s+", "", str(grade_v).strip())
            if not grade_s or grade_s in KINDERGARTEN:
                continue

            m = re.search(r"\d+", grade_s)
            if not m:
                continue

            g = int(m.group(0))
            if g != 1:
                extra.add(g)

        return sorted(extra)

    except Exception:
        return []

    finally:
        if wb is not None:
            wb.close()

# =========================
# L4. scan_pipeline
# =========================
def load_preview_rows(
    xlsx_path: Path,
    kind: str,
    header_row: int,
    data_start_row: int,
    limit: int = 100,
    sheet_name: str = "",
) -> List[List[str]]:
    """
    UI에서 미리보기 요청 시 호출.

    종료 판단과 중간 빈 행 보존은 스캔 코어와 같은 기준을 사용한다.
    즉 No가 아니라 kind별 필수 열을 기준으로 본다.
    """
    try:
        wb = safe_load_workbook(Path(xlsx_path), data_only=True)
        try:
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = get_first_sheet_with_warning(wb, Path(xlsx_path).name)

            kind_alias = {
                '신입생': 'freshmen',
                '전입생': 'transfer_in',
                '전출생': 'transfer_out',
                '교직원': 'teachers',
                '교사': 'teachers',
                '재학생 명단': 'compare',
            }
            norm_kind = kind_alias.get(kind, kind)

            compare_slots = {}
            if norm_kind == 'compare':
                from core.scan_diff import COMPARE_HEADER_SLOTS as _COMPARE_HEADER_SLOTS
                compare_slots = _COMPARE_HEADER_SLOTS

            kind_slots = {
                'freshmen': (FRESHMEN_HEADER_SLOTS, ['grade', 'class', 'name']),
                'transfer_in': (TRANSFER_HEADER_SLOTS, ['grade', 'class', 'name']),
                'transfer_out': (WITHDRAW_HEADER_SLOTS, ['grade', 'class', 'name']),
                'teachers': (TEACHER_HEADER_SLOTS, ['name']),
                'compare': (compare_slots, ['grade', 'class', 'name']),
            }
            header_slots, required_keys = kind_slots.get(norm_kind, ({}, []))
            required_cols: Dict[str, int] = {}
            if header_slots and header_row:
                slot_cols = _build_header_slot_map(ws, header_row, header_slots)
                required_cols = {
                    key: slot_cols.get(key)
                    for key in required_keys
                    if slot_cols.get(key) is not None
                }

            return _sheet_preview_rows(
                ws,
                data_start_row,
                required_cols=required_cols,
                limit=limit,
            )
        finally:
            wb.close()
    except Exception:
        return []



def scan_pipeline(
    work_root: Path,
    school_name: str,
    school_start_date: date,
    work_date: date,
    roster_basis_date: Optional[date] = None,
    roster_xlsx: Optional[Path] = None,
    col_map: Optional[Dict[str, Any]] = None,
    school_kind_override: Optional[str] = None,
) -> ScanResult:
    logs: List[str] = []

    def log(msg: str):
        from datetime import datetime as _dt
        entry = f"[{_dt.now().strftime('%H:%M:%S')}] {msg}"
        logs.append(entry)
        print("[SCAN]", entry, flush=True)

    work_root = Path(work_root).resolve()
    dirs = get_project_dirs(work_root)

    school_name = (school_name or "").strip()
    year_str = str(school_start_date.year).strip()
    year_int = school_start_date.year

    sr = ScanResult(
        ok=False,
        logs=logs,
        school_name=school_name,
        year_str=year_str,
        year_int=year_int,
        project_root=work_root,
        input_dir=Path("."),
        output_dir=Path("."),
        template_register=None,
        template_notice=None,
        roster_xlsx_path=None,
        freshmen_file=None,
        teacher_file=None,
        transfer_file=None,
        withdraw_file=None,
        need_roster=False,
        roster_path=None,
        roster_year=None,
        roster_info=None,
        needs_open_date=False,
        missing_fields=[],
        can_execute=False,
        can_execute_after_input=False,
    )
    school_profile = school_profile_from_name(school_name)
    sr.school_profile_mode = str(school_profile.get("mode", "single") or "single")
    sr.school_kind_needs_choice = bool(school_profile.get("needs_user_choice", False)) or school_profile.get("mode") == "needs_user_choice"
    sr.grade_rule_max_grade = int(school_profile.get("grade_rule_max_grade", 6) or 6)

    try:
        if not school_name:
            raise ValueError("[ERROR] 학교명을 입력해 주세요.")
        year_int = int(year_str)

        import time as _time
        _t0 = _time.perf_counter()
        def _tick(label: str):
            elapsed = _time.perf_counter() - _t0
            log(f"[TIMER] {label}: {elapsed:.2f}s")

        # 명단 xlsx 기반 학교 존재 확인 (DB 폴더 불필요)
        if roster_xlsx:
            _roster_path = Path(roster_xlsx)
            if _roster_path.suffix.lower() == ".xls":
                sr.events.append(roster_xls_format())
                raise ValueError(
                    f"[ERROR] 명단 파일이 .xls 형식입니다: {_roster_path.name}\n"
                    ".xlsx 형식으로 변환한 뒤 다시 설정해 주세요."
                )
            ensure_xlsx_only(_roster_path)
            if not _xlsx_school_exists(_roster_path, school_name, col_map):
                sr.events.append(school_not_in_roster(school_name))
                raise ValueError(
                    f"[ERROR] 명단 파일에서 '{school_name}' 학교를 찾을 수 없습니다. "
                    f"(파일: {_roster_path.name})"
                )
            sr.roster_xlsx_path = _roster_path
            log(f"[INFO] 명단 파일 검증 통과: {_roster_path.name}")
        else:
            log("[WARN] 명단 파일이 지정되지 않았습니다. 학교 존재 확인을 건너뜁니다.")
            sr.roster_xlsx_path = None
        _tick("명단 파일 검증")

        # 학교 구분 / 예외 프로필 감지
        # school_kind_needs_choice가 True면 이후 모든 검사를 건너뜀 —
        # 학교 구분 선택 전까지는 명부/파일 구조 판단 자체가 불가능하기 때문.
        _kind_full, _ = school_kind_from_name(school_name)
        if sr.school_kind_needs_choice:
            if school_kind_override:
                log(f"[INFO] 학교 구분 수동 선택: {school_kind_override}")
                sr.school_kind_needs_choice = False
            else:
                log("[WARN] 학교 구분을 자동으로 판별하지 못했습니다.")
                log("[WARN] 학교 구분을 직접 선택한 뒤 다시 실행해 주세요.")
                sr.events.append(school_kind_unknown())
                sr.ok = False
                return sr
        elif sr.school_profile_mode == "mixed":
            log(f"[INFO] 초중 통합 학교로 감지되었습니다. 학년도 규칙은 1~{sr.grade_rule_max_grade}학년까지 표시합니다.")
        elif not _kind_full:
            log("[WARN] 학교 구분을 자동으로 판별하지 못했습니다.")
            log("[WARN] 학교 구분을 직접 선택한 뒤 다시 실행해 주세요.")

        root_dir = dirs["SCHOOL_ROOT"]

        kw = (school_name or "").strip()
        if not kw:
            raise ValueError("[ERROR] 학교명을 입력해 주세요.")

        matches = [
            p
            for p in root_dir.iterdir()
            if p.is_dir() and text_contains(p.name, kw)
        ]

        if not matches:
            sr.events.append(school_folder_not_found(school_name))
            raise ValueError(
                f"[ERROR] '{school_name}' 학교 폴더를 찾을 수 없습니다."
            )

        if len(matches) > 1:
            sr.events.append(school_folder_ambiguous(school_name, [p.name for p in matches]))
            raise ValueError(
                f"[ERROR] '{school_name}' 학교 폴더 후보가 여러 개입니다: "
                + ", ".join(p.name for p in matches)
            )

        school_dir = matches[0]
        log(f"[INFO] 학교 폴더 매칭: {school_dir.name}")
        _tick("학교 폴더 매칭")

        input_dir = school_dir
        output_dir = school_dir / "작업"

        sr.input_dir = input_dir
        sr.output_dir = output_dir

        try:
            file_list = [p.name for p in input_dir.iterdir() if p.is_file()]
            log(f"[DEBUG] input files: {file_list}")
        except Exception as e:
            log(f"[WARN] 학교 폴더 파일 목록 조회 중 오류: {e}")

        try:
            freshmen_file = find_single_input_file(input_dir, FRESHMEN_KEYWORDS, file_key='freshmen', kind_label='신입생')
        except ValueError as _e:
            _es = str(_e)
            if "__DUPLICATE__" in _es: sr.events.append(duplicate_input_file("freshmen", "신입생"))
            elif "__XLS_FORMAT__" in _es: sr.events.append(input_xls_format([]))
            raise
        try:
            teacher_file = find_single_input_file(input_dir, TEACHER_KEYWORDS, file_key='teachers', kind_label='교직원')
        except ValueError as _e:
            _es = str(_e)
            if "__DUPLICATE__" in _es: sr.events.append(duplicate_input_file("teachers", "교직원"))
            elif "__XLS_FORMAT__" in _es: sr.events.append(input_xls_format([]))
            raise
        try:
            transfer_file = find_single_input_file(input_dir, TRANSFER_KEYWORDS, file_key='transfer_in', kind_label='전입생')
        except ValueError as _e:
            _es = str(_e)
            if "__DUPLICATE__" in _es: sr.events.append(duplicate_input_file("transfer_in", "전입생"))
            elif "__XLS_FORMAT__" in _es: sr.events.append(input_xls_format([]))
            raise
        try:
            withdraw_file = find_single_input_file(input_dir, WITHDRAW_KEYWORDS, file_key='transfer_out', kind_label='전출생')
        except ValueError as _e:
            _es = str(_e)
            if "__DUPLICATE__" in _es: sr.events.append(duplicate_input_file("transfer_out", "전출생"))
            elif "__XLS_FORMAT__" in _es: sr.events.append(input_xls_format([]))
            raise

        warn_if_multi_sheet(freshmen_file, logs, "신입생")
        warn_if_multi_sheet(teacher_file,  logs, "교사")
        warn_if_multi_sheet(transfer_file, logs, "전입생")
        warn_if_multi_sheet(withdraw_file, logs, "전출생")
        # 시트 수 직접 체크 — multiple_sheets 이벤트 생성
        for _fpath, _fkey, _flabel in [
            (freshmen_file, "freshmen",    "신입생"),
            (teacher_file,  "teachers",    "교사"),
            (transfer_file, "transfer_in", "전입생"),
            (withdraw_file, "transfer_out","전출생"),
        ]:
            if _fpath:
                try:
                    _wb_tmp = safe_load_workbook(_fpath, data_only=True, read_only=True)
                    if len(_wb_tmp.worksheets) > 1:
                        sr.events.append(multiple_sheets(_fkey, _flabel))
                    _wb_tmp.close()
                except Exception:
                    pass

        if not any([freshmen_file, teacher_file, transfer_file, withdraw_file]):
            sr.events.append(no_input_files())
            raise ValueError(
                "[ERROR] 신입생/전입생/전출생/교사 파일을 하나도 찾을 수 없습니다. "
                "학교 폴더 안에 해당 키워드가 포함된 xlsx 파일이 있는지 확인해 주세요."
            )

        sr.freshmen_file = freshmen_file
        sr.teacher_file = teacher_file
        sr.transfer_file = transfer_file
        sr.withdraw_file = withdraw_file

        log(f"[INFO] 신입생: {freshmen_file.name}" if freshmen_file else "[INFO] 신입생 파일 없음")
        log(f"[INFO] 교사: {teacher_file.name}" if teacher_file else "[INFO] 교사 파일 없음")
        log(f"[INFO] 전입생: {transfer_file.name}" if transfer_file else "[INFO] 전입생 파일 없음")
        log(f"[INFO] 전출생: {withdraw_file.name}" if withdraw_file else "[INFO] 전출생 파일 없음")

        template_register = choose_template_register(dirs["TEMPLATES"], year_str)
        sr.template_register = template_register
        log(f"[INFO] 양식(등록): {template_register.name}")

        template_notice = choose_template_notice(dirs["TEMPLATES"], year_str)
        sr.template_notice = template_notice
        log(f"[INFO] 양식(안내): {template_notice.name}")

        # 입력 파일 구조 자동 감지 + 경고 수집
        def run_structure_check(
            file_path: Optional[Path],
            kind_key: str,
            kind_label: str,
            header_slots: Dict[str, List[str]],
            required_keys: List[str],
            allow_blank_class_for_kindergarten: bool = False,
        ) -> Optional[Dict[str, Any]]:
            if not file_path:
                return None

            _t_file = _time.perf_counter()
            try:
                layout = detect_input_layout(file_path, kind_key)
                wb = layout["_wb"]
                ws = layout["_ws"]
            except Exception as _layout_err:
                _err_msg = str(_layout_err)
                if "데이터 시작 행" in _err_msg:
                    log(f"[ERROR] {kind_label} 데이터 시작 행을 찾을 수 없습니다.")
                    sr.events.append(missing_data_start(kind_key, kind_label))
                    _warn_text = "데이터 시작 행을 찾을 수 없습니다."
                else:
                    log(f"[ERROR] {kind_label} 파일에서 헤더를 찾을 수 없습니다.")
                    sr.events.append(missing_header(kind_key, kind_label))
                    _warn_text = "헤더를 찾을 수 없습니다."
                return {
                    "file_name": file_path.name,
                    "file_path": str(file_path),
                    "sheet_name": "",
                    "header_row": None,
                    "data_start_row": None,
                    "warning": _warn_text,
                    "headers": [],
                    "rows": [],
                    "issue_rows": [],
                    "severity": "error",
                }

            try:
                log(
                    f"[INFO] {kind_label} 자동감지 | "
                    f"헤더={layout['header_row']} | "
                    f"예시={layout['example_rows']} | "
                    f"시작={layout['data_start_row']}"
                )

                slot_cols = _build_header_slot_map(ws, layout["header_row"], header_slots)
                required_cols = {
                    key: slot_cols.get(key)
                    for key in required_keys
                    if slot_cols.get(key) is not None
                }
                # 필수열 중 매핑 못 찾은 것 → missing_required_col 이벤트
                for _rk in required_keys:
                    if slot_cols.get(_rk) is None:
                        _col_label = {"grade": "학년", "class": "반", "name": "이름"}.get(_rk, _rk)
                        sr.events.append(missing_required_col(kind_key, kind_label, _col_label))

                log(f"[DEBUG] {kind_label} 구조 검증 시작 (max_row={ws.max_row})")
                issues, issue_rows, _evts, _marks = validate_input_sheet_structure(
                    ws=ws,
                    kind=kind_label,
                    header_row=layout["header_row"],
                    data_start_row=layout["data_start_row"],
                    required_cols=required_cols,
                    allow_blank_class_for_kindergarten=allow_blank_class_for_kindergarten,
                    file_key=kind_key,
                )
                sr.events.extend(_evts)
                sr.row_marks.extend(_marks)
                log(f"[DEBUG] {kind_label} 구조 검증 완료")

                for msg in issues:
                    log(msg)

                headers = _sheet_headers(ws, layout["header_row"])
                log(f"[TIMER] {kind_label} 구조 분석: {_time.perf_counter() - _t_file:.2f}s")

                data_start = layout["data_start_row"]
                issue_row_idxs = [r for r in issue_rows if r >= data_start]

                return {
                    "file_name": file_path.name,
                    "file_path": str(file_path),
                    "sheet_name": ws.title,
                    "header_row": layout["header_row"],
                    "data_start_row": data_start,
                    "warning": "\n".join(issues) if issues else "",
                    "headers": headers,
                    "rows": [],
                    "issue_rows": issue_row_idxs,
                    "severity": "warn" if issues else "ok",
                }
            finally:
                wb.close()

        # 명부 필요 여부를 먼저 판단 — 명부 없음이 파일 구조 오류보다 우선순위가 높으므로
        # 파일 구조 검사(run_structure_check)는 명부 확인 완료 후에 실행한다.
        need_roster_by_transfer = bool(transfer_file)
        need_roster_by_withdraw = bool(withdraw_file)
        need_roster_by_freshmen = False

        if freshmen_file:
            try:
                need_roster_by_freshmen = freshmen_need_roster(
                    xlsx_path=freshmen_file,
                    input_year=year_int,
                    school_name=school_name,
                )
            except Exception as e:
                import traceback
                log(f"[DEBUG] {traceback.format_exc()}")
                raise

        need_roster = (
            need_roster_by_transfer
            or need_roster_by_withdraw
            or need_roster_by_freshmen
        )
        sr.need_roster = need_roster

        if need_roster_by_freshmen:
            extra = freshmen_extra_grades(freshmen_file, year_int, school_name)
            grade_str = ", ".join(f"{g}학년" for g in extra) if extra else "1학년 외"
            log(f"[DEBUG] 신입생 파일에 {grade_str}이(가) 포함되어 있습니다.")
            from core.events import freshmen_extra_grades_info as _feg_info
            sr.events.append(_feg_info(extra))
        elif need_roster_by_transfer or need_roster_by_withdraw:
            log("[DEBUG] 전입/전출 파일이 있어 학생명부가 필요합니다.")
        else:
            log("[DEBUG] 학생명부가 필요하지 않아 로드를 스킵합니다.")


        if need_roster:
            roster_wb = None
            roster_ws = None
            roster_path = None
            _tick("명부 로드 시작")

            try:
                try:
                    roster_wb, roster_ws, roster_path, roster_year = load_roster_sheet(dirs, school_name)
                    sr.roster_path = roster_path
                    sr.roster_year = roster_year
                    log(f"[INFO] 학생명부: {roster_path.name}")
                    _tick("명부 파일 로드")
                except ValueError as roster_err:
                    if need_roster_by_transfer or need_roster_by_withdraw:
                        missing_for = []
                        if need_roster_by_transfer:
                            missing_for.append("전입")
                        if need_roster_by_withdraw:
                            missing_for.append("전출")
                        joined = "/".join(missing_for)
                        sr.events.append(roster_not_found(reason=joined))
                        raise ValueError(
                            "[ERROR] 학생명부 파일이 없습니다. "
                            f"{joined} 처리는 학생명부 파일 없이는 진행할 수 없습니다."
                        ) from roster_err
                    else:
                        # 신입생 타학년만 있는 경우 — 명부 없이 학년도 직접 입력으로 대체 가능
                        from core.events import freshmen_no_roster_manual as _fnrm
                        sr.events.append(_fnrm())
                        sr.roster_info = None
                        log("[INFO] 신입생 타학년 포함 — 명부 추가 또는 학년도 직접 입력 후 실행 가능")

                if roster_ws is not None and roster_path is not None:
                    try:
                        sr.roster = {
                            "file_name": roster_path.name,
                            "file_path": str(roster_path),
                            "sheet_name": roster_ws.title,
                            "header_row": 1,
                            "data_start_row": 2,
                            "warning": "",
                            "headers": _sheet_headers(roster_ws, 1),
                            "rows": [],   # 미리보기는 UI에서 요청 시 별도 로드
                            "issue_rows": [],
                        }
                    except Exception as e:
                        import traceback
                        log(f"[WARN] 학생명부 미리보기 생성 중 오류: {e}")
                        log(f"[DEBUG] {traceback.format_exc()}")

                    try:
                        modified_date = datetime.fromtimestamp(roster_path.stat().st_mtime).date()
                        auto_basis = modified_date

                        if roster_basis_date is not None and roster_basis_date != auto_basis:
                            sr.roster_basis_date = roster_basis_date
                            log(
                                f"[INFO] 학생명부 마지막 수정일은 {auto_basis.isoformat()} 이지만, "
                                f"사용자가 명부 기준일을 {roster_basis_date.isoformat()} 로 수정했습니다."
                            )
                        else:
                            sr.roster_basis_date = auto_basis
                            log(
                                f"[INFO] 학생명부 마지막 수정일({auto_basis.isoformat()})을 "
                                "명부 기준일로 자동 감지했습니다."
                            )

                        if sr.roster_basis_date != work_date:
                            sr.roster_date_mismatch = True
                            log(
                                "[DEBUG] 작업일과 명부 기준일이 다릅니다. "
                                f"(작업일={work_date.isoformat()}, 명부 기준일={sr.roster_basis_date.isoformat()})"
                            )
                            # roster_date_mismatch는 모달로만 처리 — events 카드에 표시 안 함

                    except Exception as e:
                        import traceback
                        log(f"[WARN] 학생명부 파일 수정일 조회 중 오류: {e}")
                        log(f"[DEBUG] {traceback.format_exc()}")

                    try:
                        roster_info = analyze_roster_once(roster_ws, input_year=year_int)
                        sr.roster_info = roster_info
                        _tick("명부 분석(analyze_roster_once)")
                    except Exception as e:
                        import traceback
                        log(f"[WARN] 학생명부 분석 중 오류가 발생했습니다: {e}")
                        log(f"[DEBUG] {traceback.format_exc()}")
                        sr.roster_info = None

                    try:
                        roster_info = sr.roster_info
                        basis_date = roster_basis_date or sr.roster_basis_date or work_date
                        sr.roster_basis_date = basis_date

                        # 명부 기준일 ↔ 개학일 비교로 roster_time/ref_grade_shift 확정
                        if basis_date < school_start_date:
                            roster_time = "last_year"
                            ref_shift = -1
                        else:
                            roster_time = "this_year"
                            ref_shift = 0

                        if roster_info is not None:
                            roster_info.roster_time     = roster_time
                            roster_info.ref_grade_shift = ref_shift
                            sr.roster_info = roster_info

                        log(f"[DEBUG] 명부 기준일 기반 학년도: {'작년' if roster_time == 'last_year' else '올해'}")
                        log(
                            "[INFO] 명부 기준일/개학일 기준으로 "
                            f"'{'작년' if roster_time == 'last_year' else '올해'} 학년도 명부'로 간주합니다."
                        )
                    except Exception as e:
                        import traceback
                        log(f"[WARN] 학생명부 학년도 추정 중 오류가 발생했습니다: {e}")
                        log(f"[DEBUG] {traceback.format_exc()}")

            finally:
                if roster_wb is not None:
                    roster_wb.close()
        else:
            log("[DEBUG] 전입/전출/1학년 외 신입생 케이스가 없어 학생명부 로드를 건너뜁니다.")

        needs_open_date = bool(withdraw_file)
        sr.needs_open_date = needs_open_date
        if needs_open_date:
            log("[INFO] 전출생 파일 감지 → 개학일(퇴원일자 계산용) 입력 필요")
            if not school_start_date:
                sr.events.append(open_date_missing())
        else:
            log("[INFO] 전출생 파일 없음 → 개학일 입력 불필요")

        missing_fields: List[str] = []

        if sr.template_register is None:
            missing_fields.append("등록 템플릿")
        if sr.template_notice is None:
            missing_fields.append("안내 템플릿")

        # 명부 확인 완료 후 파일 구조 검사 실행
        # 우선순위: 학교구분 판별불가 > 명부 없음 > 파일 구조 오류
        # 명부 없음이면 이미 raise로 빠져나왔으므로 여기까지 온 경우만 구조 검사함
        sr.freshmen = run_structure_check(
            file_path=freshmen_file,
            kind_key="freshmen",
            kind_label="신입생",
            header_slots=FRESHMEN_HEADER_SLOTS,
            required_keys=["grade", "class", "name"],
            allow_blank_class_for_kindergarten=True,
        )

        sr.transfer_in = run_structure_check(
            file_path=transfer_file,
            kind_key="transfer",
            kind_label="전입생",
            header_slots=TRANSFER_HEADER_SLOTS,
            required_keys=["grade", "class", "name"],
            allow_blank_class_for_kindergarten=True,
        )

        sr.transfer_out = run_structure_check(
            file_path=withdraw_file,
            kind_key="withdraw",
            kind_label="전출생",
            header_slots=WITHDRAW_HEADER_SLOTS,
            required_keys=["grade", "class", "name"],
            allow_blank_class_for_kindergarten=True,
        )

        sr.teachers = run_structure_check(
            file_path=teacher_file,
            kind_key="teacher",
            kind_label="교사",
            header_slots=TEACHER_HEADER_SLOTS,
            required_keys=["name"],
        )

        if (
            freshmen_file and transfer_file
            and sr.freshmen and sr.transfer_in
            and sr.freshmen.get("severity") != "error"
            and sr.transfer_in.get("severity") != "error"
        ):
            try:
                from core.run_main import read_freshmen_rows, read_transfer_rows
                from core.events import freshmen_transfer_same_student_scan

                freshmen_rows, _, _ = read_freshmen_rows(
                    freshmen_file,
                    input_year=year_int,
                    header_row=sr.freshmen.get("header_row"),
                    data_start_row=sr.freshmen.get("data_start_row"),
                    roster_info=sr.roster_info,
                    school_name=school_name,
                )
                transfer_rows, _, _ = read_transfer_rows(
                    transfer_file,
                    header_row=sr.transfer_in.get("header_row"),
                    data_start_row=sr.transfer_in.get("data_start_row"),
                )

                def _norm_class(value: Any) -> str:
                    if value is None:
                        return ""
                    return re.sub(r"\s+", "", str(value).strip())

                freshmen_keys = {
                    (
                        int(row.get("grade", 0) or 0),
                        _norm_class(row.get("class", "")),
                        normalize_name_key(row.get("name", "")),
                    )
                    for row in freshmen_rows
                    if normalize_name_key(row.get("name", ""))
                }

                seen_keys = set()
                for row in transfer_rows:
                    key = (
                        int(row.get("grade", 0) or 0),
                        _norm_class(row.get("class", "")),
                        normalize_name_key(row.get("name", "")),
                    )
                    if key in freshmen_keys and key not in seen_keys:
                        seen_keys.add(key)
                        sr.events.append(
                            freshmen_transfer_same_student_scan(
                                grade=int(row.get("grade", 0) or 0),
                                class_=str(row.get("class", "")).strip(),
                                name=str(row.get("name", "")).strip(),
                            )
                        )
            except Exception:
                pass

        # 교직원 파일이 있고 관리용 ID 신청자가 0건이면 스캔 단계에서 경고
        if teacher_file and sr.teachers and (sr.teachers.get("severity") != "error"):
            try:
                from core.run_main import read_teacher_rows
                _t_rows, _, _ = read_teacher_rows(
                    teacher_file,
                    header_row=sr.teachers.get("header_row"),
                    data_start_row=sr.teachers.get("data_start_row"),
                )
                if _t_rows and not any(r.get("admin_apply") for r in _t_rows):
                    from core.events import no_teacher_id_request as _evt_no_tid
                    sr.events.append(_evt_no_tid())
            except Exception:
                pass

        if any(
            (meta or {}).get("severity") == "error"
            for meta in [sr.freshmen, sr.transfer_in, sr.transfer_out, sr.teachers]
            if meta is not None
        ):
            missing_fields.append("입력 파일 헤더/구조 확인")

        # 신입생 타학년만 있고 명부 없는 경우 → 학년도 수동 입력으로 실행 가능
        sr.can_execute_after_input = (
            need_roster_by_freshmen
            and not need_roster_by_transfer
            and not need_roster_by_withdraw
            and sr.roster_info is None
        )

        roster_ok = (not sr.need_roster) or (sr.roster_info is not None)
        if not roster_ok and not sr.can_execute_after_input:
            # 수동 입력 불가한 케이스만 오류 처리
            missing_fields.append("학생명부")
            if need_roster_by_transfer and need_roster_by_withdraw:
                log("[ERROR] 전입/전출 처리를 위해 학생명부가 필요합니다. 학생명부를 추가한 뒤 다시 스캔해 주세요.")
            elif need_roster_by_transfer:
                log("[ERROR] 전입생 처리를 위해 학생명부가 필요합니다. 학생명부를 추가한 뒤 다시 스캔해 주세요.")
            elif need_roster_by_withdraw:
                log("[ERROR] 전출생 처리를 위해 학생명부가 필요합니다. 학생명부를 추가한 뒤 다시 스캔해 주세요.")
            else:
                log("[ERROR] 학년도 규칙이 필요한 경우에도 학생명부가 필요합니다. 학생명부를 추가한 뒤 다시 스캔해 주세요.")

        sr.missing_fields = missing_fields
        sr.needs_open_date = bool(sr.withdraw_file)

        if sr.can_execute_after_input:
            # grade_year_map 입력 전까지는 실행 불가 — applyManualGradeReady에서 활성화
            sr.can_execute = False
        else:
            sr.can_execute = len(sr.missing_fields) == 0

        sr.ok = True
        _tick("스캔 완료 (전체)")
        log("[DONE] 스캔 완료")
        return sr

    except Exception as e:
        import traceback
        # ValueError는 사용자에게 설명 가능한 오류 — 메시지 그대로
        # 그 외는 예상치 못한 예외 — DEBUG로 traceback 보존
        if not isinstance(e, ValueError):
            log(f"[DEBUG] {traceback.format_exc()}")
        log(f"[ERROR] {e}")
        sr.ok = False
        return sr

# =========================
# L5. 작업 루트 점검
# =========================
def ensure_work_root_scaffold(work_root: Path) -> List[str]:
    """
    작업 폴더 하위에 resources/templates/notices 폴더 구조를 생성한다.
    이미 있는 폴더는 건드리지 않는다 (exist_ok=True).

    반환: 이번 호출에서 새로 생성된 폴더 이름 목록 (빈 리스트면 이미 모두 존재)
    """
    work_root = Path(work_root).resolve()
    dirs = get_project_dirs(work_root)
    scaffolded: List[str] = []
    for key, label in [
        ("RESOURCES_ROOT", "resources"),
        ("TEMPLATES",      "resources/templates"),
        ("NOTICES",        "resources/notices"),
    ]:
        folder = dirs[key]
        if not folder.exists():
            try:
                folder.mkdir(parents=True, exist_ok=True)
                scaffolded.append(label)
            except Exception:
                pass
    return scaffolded


def scan_work_root(work_root: Path) -> Dict[str, Any]:
    """
    앱 시작 시 1회 호출하여 resources 폴더 구조를 점검한다.
    engine.py의 inspect_work_root()가 이 함수를 호출한다.
    순수 조회만 수행한다 — 폴더 생성 등 부작용 없음.

    반환 dict 키:
      ok                — 전체 이상 없음 여부
      errors            — 오류 메시지 목록
      message           — ok일 때 성공 메시지
      school_folders    — 작업 루트 내 학교 폴더 이름 목록
      notice_titles     — notices/*.txt 파일명(stem) 목록
      format_ok         — templates 폴더 정상 여부
      errors_format     — templates 관련 오류 목록
      register_template — 등록 템플릿 경로 (없으면 None)
      notice_template   — 안내 템플릿 경로 (없으면 None)
    """
    work_root = work_root.resolve()
    dirs = get_project_dirs(work_root)

    errors: List[str] = []

    res_root = dirs["RESOURCES_ROOT"].resolve()
    school_folders = sorted(
        p.name for p in work_root.iterdir()
        if p.is_dir() and p.resolve() != res_root and not p.name.startswith(".")
    )

    format_ok = False; errors_format: List[str] = []
    register_template: Optional[Path] = None; notice_template: Optional[Path] = None
    tpl_dir = dirs["TEMPLATES"]
    if not tpl_dir.exists():
        errors_format.append("[ERROR] resources/templates 폴더가 없습니다.")
    else:
        import unicodedata as _ud
        reg_f = [p for p in tpl_dir.glob("*.xlsx") if "등록" in _ud.normalize('NFC', p.stem) and not p.name.startswith("~$")]
        ntc_f = [p for p in tpl_dir.glob("*.xlsx") if "안내" in _ud.normalize('NFC', p.stem) and not p.name.startswith("~$")]
        if len(reg_f) != 1:
            errors_format.append("[ERROR] templates 폴더에 '등록' 템플릿 파일이 정확히 1개 있어야 합니다.")
        else: register_template = reg_f[0]
        if len(ntc_f) != 1:
            errors_format.append("[ERROR] templates 폴더에 '안내' 템플릿 파일이 정확히 1개 있어야 합니다.")
        else: notice_template = ntc_f[0]
        if not errors_format: format_ok = True

    notice_titles: List[str] = []
    notice_dir = dirs["NOTICES"]
    if not notice_dir.exists():
        errors.append("[ERROR] resources/notices 폴더가 없습니다.")
    else:
        txt_files = [p for p in notice_dir.glob("*.txt") if p.is_file()]
        if not txt_files: errors.append("[ERROR] notices 폴더에 .txt 파일이 없습니다.")
        else: notice_titles = sorted({p.stem.strip() for p in txt_files})

    errors.extend(errors_format)
    ok = not errors

    return {
        "ok": ok,
        "errors": errors,
        "message": "[INFO] resources(templates/notices)가 정상적으로 준비되었습니다." if ok else "",
        "school_folders": school_folders,
        "notice_titles": notice_titles,
        "format_ok": format_ok, "errors_format": errors_format,
        "register_template": register_template, "notice_template": notice_template,
    }
