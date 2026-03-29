# core/scan_diff.py
"""
비교(diff) 파이프라인의 스캔 전용 모듈.

책임 범위:
  - 학생명렬표 / 비교용 입력 파일 탐색
  - 헤더 행 / 데이터 시작 행 / 주요 컬럼 자동 감지
  - 비교 대상 행 읽기
  - 명부 비교 전 필요한 구조 검증
  - scan_diff_pipeline() 실행 결과 반환

이 모듈은 diff 작업의 사전 분석(scan)만 담당.

공개 API:
  DiffScanResult
  scan_diff_pipeline(work_root, school_name, target_year, school_start_date, work_date, roster_basis_date)
  -> DiffScanResult

  build_diff_rows(roster_rows, compare_rows) -> Dict[str, List]
  read_roster_compare_rows(roster_ws, target_grades, ref_grade_shift) -> List[Dict]
  read_compare_rows(xlsx_path, header_row, data_start_row) -> List[Dict]
  collect_text_only_classes_from_roster(roster_ws, target_grades, ref_grade_shift) -> List[str]
  TARGET_GRADES
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils import text_contains, normalize_text
from core.events import (
    CoreEvent, RowMark,
    school_folder_not_found, school_folder_ambiguous,
    school_not_in_roster, roster_not_found, roster_date_mismatch,
    missing_header, missing_required_col,
    compare_file_not_found, compare_file_format_warn,
)
from core.common import (
    get_project_dirs,
    ensure_xlsx_only,
    safe_load_workbook,
    _build_header_slot_map,
    _detect_header_row_generic,
    normalize_name,
    normalize_name_key,
    normalize_compare_name,
    normalize_compare_name_key,
    get_compare_name_warnings,
    load_roster_sheet,
    parse_roster_year_from_filename,
    parse_class_str,
)


# =========================
# Result types
# =========================
@dataclass
class DiffScanResult:
    ok: bool = False
    logs: List[str] = field(default_factory=list)

    school_name: str = ""
    year_str: str = ""
    year_int: int = 0

    project_root: Path = Path(".")
    input_dir: Path = Path(".")
    output_dir: Path = Path(".")

    roster_xlsx_path: Optional[Path] = None
    roster_path: Optional[Path] = None
    roster_year: Optional[int] = None

    compare_file: Optional[Path] = None
    template_transfer_in: Optional[Path] = None
    template_transfer_out: Optional[Path] = None

    compare_layout: Optional[Dict[str, Any]] = None

    missing_fields: List[str] = field(default_factory=list)
    can_execute: bool = False

    roster_date_mismatch: bool = False
    roster_basis_date: Optional[date] = None
    school_start_date: Optional[date] = None
    work_date: Optional[date] = None
    ref_grade_shift: int = 0

    # 구조화된 판정 결과 — bridge/UI는 이것만 참조
    events:    List[Any] = field(default_factory=list)  # List[CoreEvent]
    row_marks: List[Any] = field(default_factory=list)  # List[RowMark]


# =========================
# Constants
# =========================
EXCLUDED_CLASS_KEYWORDS = [
    "선생님반",
    "유치원반",
    "선생님자녀반",
    "테스트반",
]

COMPARE_FILE_KEYWORDS = [
    "명렬표",
    "명렬",
    "재학생",
    "학생명단",
    "학생 명단",
]

DIFF_TRANSFER_IN_TEMPLATE_KEYWORDS = ["전입생"]
DIFF_TRANSFER_OUT_TEMPLATE_KEYWORDS = ["전출생"]

COMPARE_HEADER_SLOTS = {
    "grade": ["학년"],
    "class": ["반", "학급"],
    "name": ["이름", "성명", "학생이름"],
}

TARGET_GRADES = set(range(2, 7))

ROSTER_HEADER_SLOTS = {
    "current_class": ["현재반"],
    "previous_class": ["이전반"],
    "name": ["학생이름", "이름", "성명"],
    "student_id": ["아이디", "ID"],
}


# =========================
# L1. Compare input / layout detection
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


def detect_example_and_data_start(ws, header_row: int, name_col: int, max_search_row: Optional[int] = None, max_col: Optional[int] = None) -> Tuple[List[int], int]:
    """메인 스캔과 동일 기준으로 예시 행과 실제 데이터 시작 행을 감지한다."""
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

    raise ValueError('[ERROR] 데이터 시작 행을 찾을 수 없습니다. 헤더 행 아래에 실제 데이터가 있는지 확인해 주세요.')

def detect_header_row_compare(ws) -> int:
    """
    비교용 재학생 명렬표 헤더 탐지.
    - 필수는 학년, 이름
    - 반은 보통 오지만 없어도 허용
    """
    try:
        return _detect_header_row_generic(
            ws,
            COMPARE_HEADER_SLOTS,
            max_search_row=15,
            max_col=10,
            min_match_slots=2,
        )
    except Exception:
        raise ValueError(
            "[ERROR] 재학생 명렬표 파일에서 헤더를 찾을 수 없습니다. "
            "'학년', '이름' 열이 있는지 확인해 주세요."
        )


def detect_compare_input_layout(xlsx_path: Path) -> Dict[str, Any]:
    ensure_xlsx_only(xlsx_path)
    wb = safe_load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb.worksheets[0]

        header_row = detect_header_row_compare(ws)
        slot_cols = _build_header_slot_map(ws, header_row, COMPARE_HEADER_SLOTS)

        name_col = slot_cols.get("name")
        if name_col is None:
            raise ValueError("[ERROR] 재학생 명렬표 헤더에서 이름 열을 찾을 수 없습니다.")

        example_rows, data_start_row = detect_example_and_data_start(
            ws,
            header_row=header_row,
            name_col=name_col,
            max_col=min(ws.max_column or 1, 10),
        )

        return {
            "header_row": header_row,
            "example_rows": example_rows,
            "data_start_row": data_start_row,
            "slot_cols": slot_cols,
        }
    finally:
        wb.close()


def validate_compare_input_rows(ws, header_row: int, data_start_row: int, slot_cols: Dict[str, int]) -> Tuple[List[str], List[int], int, List[CoreEvent], List[RowMark]]:
    """
    비교용 재학생 명단 시트 경고 검증.
    - 이름은 원문 기준으로 비교하고, 형식만 WARN 처리한다.
    - 메인 스캔 톤과 맞춰 [WARN] 문구를 누적한다.
    반환: (issues, issue_row_nums, row_count, events, row_marks)
    """
    issues: List[str] = []
    issue_row_nums: List[int] = []
    evts: List[CoreEvent] = []
    marks: List[RowMark] = []

    col_grade = slot_cols.get("grade")
    col_class = slot_cols.get("class")
    col_name = slot_cols.get("name")
    if col_grade is None or col_name is None:
        return issues, issue_row_nums, 0

    max_col = max([c for c in [col_grade, col_class, col_name] if c is not None], default=1)
    last_data_row = data_start_row - 1
    blank_streak = 0
    row_count = 0

    for i, row_tuple in enumerate(ws.iter_rows(min_row=data_start_row, max_col=max_col, values_only=True)):
        r = data_start_row + i
        grade_v = row_tuple[col_grade - 1] if col_grade - 1 < len(row_tuple) else None
        class_v = row_tuple[col_class - 1] if (col_class is not None and col_class - 1 < len(row_tuple)) else None
        name_v = row_tuple[col_name - 1] if col_name - 1 < len(row_tuple) else None

        vals = [grade_v, class_v, name_v]
        has_any = any(v is not None and str(v).strip() != "" for v in vals)
        if has_any:
            last_data_row = r
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak >= 10:
                break
            continue

        row_count += 1

        grade_s = "" if grade_v is None else str(grade_v).strip()
        name_s = "" if name_v is None else str(name_v).strip()
        if not grade_s or not name_s:
            issues.append(f"[WARN] 재학생 파일 {r}행 '학년/이름' 값을 확인해 주세요.")
            issue_row_nums.append(r)
            _e, _m = compare_file_format_warn(r, "학년/이름", "값이 비어 있습니다.")
            evts.append(_e); marks.append(_m)
            continue

        gs_norm = re.sub(r"\s+", "", grade_s)
        warn_grade = None
        if "-" in grade_s:
            warn_grade = f"학년 열에 하이픈(-) 포함 — 학년+반이 합쳐진 것 같습니다: '{grade_s}'"
        elif gs_norm.endswith("학년"):
            warn_grade = f"학년 열에 '학년' 글자 포함 — 숫자만 입력해야 합니다: '{grade_s}'"
        else:
            try:
                float(grade_s)
            except ValueError:
                warn_grade = f"학년 열에 숫자가 아닌 값이 있습니다: '{grade_s}'"
        if warn_grade:
            issues.append(f"[WARN] 재학생 파일 {r}행 '학년' 열 — {warn_grade}")
            issue_row_nums.append(r)
            _e, _m = compare_file_format_warn(r, "학년", warn_grade)
            evts.append(_e); marks.append(_m)

        for msg in get_compare_name_warnings(name_v):
            issues.append(f"[WARN] 재학생 파일 {r}행 '{normalize_compare_name(name_v)}' — {msg}")
            issue_row_nums.append(r)
            _e, _m = compare_file_format_warn(r, "이름", msg)
            evts.append(_e); marks.append(_m)

    return issues, sorted(set(issue_row_nums)), row_count, evts, marks


def find_diff_templates(template_dir: Path) -> Tuple[Optional[Path], Optional[Path], List[str]]:
    """
    resources/templates 폴더에서
    - 파일명에 '전입생' 포함된 템플릿 1개
    - 파일명에 '전출생' 포함된 템플릿 1개
    를 찾는다.
    """
    template_dir = Path(template_dir).resolve()
    if not template_dir.exists():
        return None, None, [f"[ERROR] templates 폴더를 찾을 수 없습니다."]

    xlsx_files = [
        p for p in template_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
    ]
    if not xlsx_files:
        return None, None, [f"[ERROR] templates 폴더에 xlsx 파일이 없습니다."]

    transfer_in_files  = [p for p in xlsx_files if "전입생" in p.stem]
    transfer_out_files = [p for p in xlsx_files if "전출생" in p.stem]

    errors: List[str] = []

    if len(transfer_in_files) == 0:
        errors.append("[ERROR] templates 폴더에서 '전입생' 템플릿을 찾을 수 없습니다.")
    elif len(transfer_in_files) > 1:
        errors.append("[ERROR] templates 폴더에 '전입생' 템플릿이 여러 개 있습니다.")

    if len(transfer_out_files) == 0:
        errors.append("[ERROR] templates 폴더에서 '전출생' 템플릿을 찾을 수 없습니다.")
    elif len(transfer_out_files) > 1:
        errors.append("[ERROR] templates 폴더에 '전출생' 템플릿이 여러 개 있습니다.")

    if errors:
        return None, None, errors

    return transfer_in_files[0], transfer_out_files[0], []


def find_compare_file(input_dir: Path) -> Optional[Path]:
    """
    학교 폴더 안 xlsx 중에서 비교용 재학생 명렬표 파일 1개를 찾는다.
    1차: 파일명 키워드 후보 수집
    2차: 헤더 구조 검증
    """
    if not input_dir.exists():
        return None

    xlsx_files = [
        p for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
    ]
    if not xlsx_files:
        return None

    keyword_candidates = [
        p for p in xlsx_files
        if any(text_contains(p.name, kw) for kw in COMPARE_FILE_KEYWORDS)
    ]

    candidates = keyword_candidates if keyword_candidates else xlsx_files

    valid: List[Path] = []
    for p in candidates:
        try:
            detect_compare_input_layout(p)
            valid.append(p)
        except Exception:
            continue

    if len(valid) == 0:
        return None
    if len(valid) > 1:
        raise ValueError(f"[ERROR] 비교용 재학생 명렬표 파일 후보가 여러 개입니다: {[p.name for p in valid]}")
    return valid[0]


# =========================
# L2. Compare domain rules
# =========================
def is_excluded_misc_class(raw: Any) -> bool:
    if raw is None:
        return False

    s = str(raw).strip()
    if not s:
        return False

    s = re.sub(r"\s+", "", s)
    return any(text_contains(s, kw) for kw in EXCLUDED_CLASS_KEYWORDS)


def parse_grade_int(raw: Any) -> Optional[int]:
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return None
    return int(m.group(0))


def normalize_class_value(raw: Any) -> str:
    """
    비교용 파일/명부 공통 반 표기 정규화.
    - 없으면 ""
    - 숫자 추출 가능하면 숫자만 반환
    - 아니면 공백만 정리한 원문 반환
    """
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""

    s = s.replace("\u3000", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", "", s)

    nums = re.findall(r"\d+", s)
    if nums:
        return str(int(nums[-1]))
    return s


def build_compare_key(grade: int, name_key: str) -> Tuple[int, str]:
    return grade, name_key


def read_compare_rows(
    xlsx_path: Path,
    header_row: Optional[int] = None,
    data_start_row: Optional[int] = None,
    slot_cols: Optional[Dict[str, int]] = None,
) -> List[Dict[str, Any]]:
    """
    재학생 명렬표 비교용 입력 읽기.
    필수: 학년, 이름
    선택: 반
    """
    ensure_xlsx_only(xlsx_path)
    wb = safe_load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb.worksheets[0]

        if header_row is None or data_start_row is None:
            layout = detect_compare_input_layout(xlsx_path)
            header_row = layout["header_row"]
            data_start_row = layout["data_start_row"]
            if slot_cols is None:
                slot_cols = layout.get("slot_cols")

        # IMPORTANT:
        # 실행 단계에서는 스캔에서 확정한 slot_cols를 그대로 재사용해야 한다.
        # 여기서 다시 헤더 슬롯을 재감지하면, 스캔 때는 잡힌 '반' 열이
        # 실행 때 누락되어 판정불가 시트의 '명단반'이 공백이 될 수 있다.
        if slot_cols is None:
            slot_cols = _build_header_slot_map(ws, header_row, COMPARE_HEADER_SLOTS)
        col_grade = slot_cols.get("grade")
        col_class = slot_cols.get("class")
        col_name  = slot_cols.get("name")

        missing: List[str] = []
        if col_grade is None:
            missing.append("학년")
        if col_name is None:
            missing.append("이름")

        if missing:
            raise ValueError(
                "[ERROR] 재학생 명렬표 헤더에서 "
                + ", ".join(missing)
                + " 열을 찾을 수 없습니다."
            )

        out: List[Dict[str, Any]] = []
        row = data_start_row

        while row <= ws.max_row:
            grade_v = ws.cell(row=row, column=col_grade).value
            class_v = ws.cell(row=row, column=col_class).value if col_class is not None else None
            name_v  = ws.cell(row=row, column=col_name).value

            vals = [grade_v, class_v, name_v]

            if all(v is None or str(v).strip() == "" for v in vals):
                row += 1
                continue

            if grade_v is None or str(grade_v).strip() == "" or name_v is None or str(name_v).strip() == "":
                raise ValueError(
                    f"[ERROR] 재학생 명렬표 {row}행에 학년/이름 중 빈 값이 있습니다."
                )

            grade_i = parse_grade_int(grade_v)
            if grade_i is None:
                raise ValueError(
                    f"[ERROR] 재학생 명렬표 {row}행에서 학년 값을 인식할 수 없습니다: {grade_v!r}"
                )

            if grade_i not in TARGET_GRADES:
                row += 1
                continue

            name_n = normalize_compare_name(name_v)
            if not name_n:
                raise ValueError(
                    f"[ERROR] 재학생 명렬표 {row}행 이름을 인식할 수 없습니다."
                )

            class_raw = "" if class_v is None else str(class_v).strip()
            class_s   = normalize_class_value(class_v)

            out.append(
                {
                    "grade": grade_i,
                    "class": class_s,
                    "class_raw": class_raw,
                    "name": name_n,
                    "name_key": normalize_compare_name_key(name_n),
                    "source_row": row,
                }
            )
            row += 1

        return out
    finally:
        wb.close()


def _detect_roster_header_row(roster_ws) -> int:
    return _detect_header_row_generic(
        roster_ws,
        ROSTER_HEADER_SLOTS,
        max_search_row=20,
        max_col=20,
        min_match_slots=3,
    )


def _get_roster_slot_cols(roster_ws) -> Tuple[int, Optional[int], Optional[int], int, Optional[int]]:
    header_row = _detect_roster_header_row(roster_ws)
    slot_cols = _build_header_slot_map(roster_ws, header_row, ROSTER_HEADER_SLOTS)

    col_now = slot_cols.get("current_class")
    col_prev = slot_cols.get("previous_class")
    col_name = slot_cols.get("name")
    col_id = slot_cols.get("student_id")

    missing: List[str] = []
    if col_now is None:
        missing.append("현재반")
    if col_prev is None:
        missing.append("이전반")
    if col_name is None:
        missing.append("학생이름")

    if missing:
        raise ValueError(
            "[ERROR] 학생명부 헤더에서 " + ", ".join(missing) + " 열을 찾을 수 없습니다."
        )

    return header_row, col_now, col_prev, col_name, col_id


def _pick_roster_class_values(now_v: Any, prev_v: Any) -> Tuple[Any, str, str, str]:
    now_raw = "" if now_v is None else str(now_v).strip()
    prev_raw = "" if prev_v is None else str(prev_v).strip()

    class_source = now_v if now_raw else prev_v
    class_raw = now_raw or prev_raw
    class_s = normalize_class_value(class_source)
    return class_source, class_raw, class_s, now_raw


def _parse_roster_compare_grade(class_value: Any, ref_grade_shift: int) -> Optional[int]:
    if class_value is None:
        return None
    parsed = parse_class_str(class_value)
    if parsed is None:
        return None
    roster_grade = parsed[0]
    return roster_grade - ref_grade_shift


def collect_text_only_classes_from_roster(
    roster_ws,
    target_grades: Optional[set] = None,
    ref_grade_shift: int = 0,
) -> List[str]:
    """
    학생명부에서 숫자가 전혀 없는 현재반명을 수집한다.
    메인 로직과 동일하게 현재반을 기본으로 보되, 현재반이 비어 있으면 이전반을 보조로 사용한다.
    """
    if target_grades is None:
        target_grades = TARGET_GRADES

    header_row, col_now, col_prev, col_name, _ = _get_roster_slot_cols(roster_ws)

    found = set()
    data_start_row = header_row + 1
    row = data_start_row

    while row <= roster_ws.max_row:
        now_v = roster_ws.cell(row=row, column=col_now).value
        prev_v = roster_ws.cell(row=row, column=col_prev).value if col_prev is not None else None
        name_v = roster_ws.cell(row=row, column=col_name).value

        vals = [now_v, prev_v, name_v]
        if all(v is None or str(v).strip() == "" for v in vals):
            row += 1
            continue
        if name_v is None or str(name_v).strip() == "":
            row += 1
            continue

        class_value, class_raw, _, _ = _pick_roster_class_values(now_v, prev_v)
        compare_grade = _parse_roster_compare_grade(class_value, ref_grade_shift)
        if compare_grade is None or compare_grade not in target_grades:
            row += 1
            continue

        if not class_raw:
            row += 1
            continue

        class_compact = re.sub(r"\s+", "", class_raw)

        if is_excluded_misc_class(class_compact):
            row += 1
            continue

        if not re.search(r"\d", class_compact):
            found.add(class_raw)

        row += 1

    return sorted(found)


def read_roster_compare_rows(
    roster_ws,
    target_grades: Optional[set] = None,
    ref_grade_shift: int = 0,
) -> List[Dict[str, Any]]:
    """
    학생명부 ws에서 비교용 (학년, 반, 이름) 목록만 읽는다.

    메인 로직 기준:
    - 반/학년 파싱은 현재반을 기본으로 사용
    - 현재반이 비어 있으면 이전반을 보조로 사용
    - 비교 키는 학년+이름만 쓰므로, 반이 비어도 행 자체를 버리지 않는다
    """
    if target_grades is None:
        target_grades = TARGET_GRADES

    header_row, col_now, col_prev, col_name, col_id = _get_roster_slot_cols(roster_ws)

    data_start_row = header_row + 1
    out: List[Dict[str, Any]] = []
    row = data_start_row

    while row <= roster_ws.max_row:
        now_v = roster_ws.cell(row=row, column=col_now).value
        prev_v = roster_ws.cell(row=row, column=col_prev).value if col_prev is not None else None
        name_v = roster_ws.cell(row=row, column=col_name).value
        id_v = roster_ws.cell(row=row, column=col_id).value if col_id is not None else None

        vals = [now_v, prev_v, name_v]
        if all(v is None or str(v).strip() == "" for v in vals):
            row += 1
            continue
        if name_v is None or str(name_v).strip() == "":
            row += 1
            continue

        class_value, class_raw, class_s, now_class_raw = _pick_roster_class_values(now_v, prev_v)
        compare_grade = _parse_roster_compare_grade(class_value, ref_grade_shift)
        if compare_grade is None or compare_grade not in target_grades:
            row += 1
            continue

        name_n = normalize_compare_name(name_v)
        if not name_n:
            row += 1
            continue

        if re.sub(r"\s+", "", class_raw) == "테스트반":
            row += 1
            continue

        out.append(
            {
                "grade": compare_grade,
                "class": class_s,
                "class_raw": class_raw,
                "now_class_raw": now_class_raw,
                "name": name_n,
                "name_key": normalize_compare_name_key(name_n),
                "student_id": "" if id_v is None else str(id_v).strip(),
                "source_row": row,
            }
        )
        row += 1

    return out


def _group_by_key(rows: List[Dict[str, Any]]) -> Dict[Tuple[int, str], List[Dict[str, Any]]]:
    grouped: Dict[Tuple[int, str], List[Dict[str, Any]]] = {}
    for r in rows:
        key = build_compare_key(r["grade"], r["name_key"])
        grouped.setdefault(key, []).append(r)
    return grouped


def build_diff_rows(
    roster_rows: List[Dict[str, Any]],
    compare_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    분류 결과:
    - matched_rows        : 양쪽 모두 있음 (정상 재학생)
    - compare_only_rows   : 학교 명단에만 있음 (신규 유입 후보)
    - roster_only_rows    : 명부에만 있음 (누락/전출 후보)
    - unresolved_rows     : 자동 판정 불가 (중복 등)
    - transfer_in_done    : 전입 자동 분류 완료
    - transfer_in_hold    : 전입 보류
    - transfer_out_done   : 전출 자동 분류 완료
    - transfer_out_hold   : 전출 보류
    """
    roster_map  = _group_by_key(roster_rows)
    compare_map = _group_by_key(compare_rows)

    all_keys = sorted(set(roster_map.keys()) | set(compare_map.keys()))

    matched_rows:      List[Dict[str, Any]] = []
    compare_only_rows: List[Dict[str, Any]] = []
    roster_only_rows:  List[Dict[str, Any]] = []
    unresolved_rows:   List[Dict[str, Any]] = []

    transfer_in_done:  List[Dict[str, Any]] = []
    transfer_in_hold:  List[Dict[str, Any]] = []
    transfer_out_done: List[Dict[str, Any]] = []
    transfer_out_hold: List[Dict[str, Any]] = []

    for key in all_keys:
        roster_group  = roster_map.get(key, [])
        compare_group = compare_map.get(key, [])

        roster_count  = len(roster_group)
        compare_count = len(compare_group)

        # 한쪽이라도 중복이면 자동 판정 제외
        if roster_count >= 2 or compare_count >= 2:
            if compare_count > roster_count:
                for r in compare_group:
                    rec = {
                        "grade": r["grade"],
                        "class": r.get("class", ""),  # 기존 정렬/호환 유지용
                        "name": r["name"],
                        "source": "명단",
                        "compare_class": r.get("class", ""),
                        "roster_class": "",
                        "hold_reason": "명단 중복 / 명부 단일로 자동 판정 불가",
                    }
                    transfer_in_hold.append(rec)
                    unresolved_rows.append(rec)

            elif roster_count > compare_count:
                for r in roster_group:
                    cls_raw = r.get("class_raw", r.get("class", ""))
                    if is_excluded_misc_class(cls_raw):
                        continue
                    rec = {
                        "grade": r["grade"],
                        "class": r.get("class", ""),  # 기존 정렬/호환 유지용
                        "name": r["name"],
                        "source": "명부",
                        "compare_class": "",
                        "roster_class": r.get("class", ""),
                        "hold_reason": "명단 단일 / 명부 중복으로 자동 판정 불가",
                    }
                    transfer_out_hold.append(rec)
                    unresolved_rows.append(rec)
            else:
                for r in compare_group:
                    rec = {
                        "grade": r["grade"],
                        "class": r.get("class", ""),  # 기존 정렬/호환 유지용
                        "name": r["name"],
                        "source": "명단",
                        "compare_class": r.get("class", ""),
                        "roster_class": "",
                        "hold_reason": "학년+이름 중복(양쪽 동수)으로 자동 판정 불가",
                    }
                    transfer_in_hold.append(rec)
                    unresolved_rows.append(rec)
                for r in roster_group:
                    cls_raw = r.get("class_raw", r.get("class", ""))
                    if is_excluded_misc_class(cls_raw):
                        continue
                    rec = {
                        "grade": r["grade"],
                        "class": r.get("class", ""),  # 기존 정렬/호환 유지용
                        "name": r["name"],
                        "source": "명부",
                        "compare_class": "",
                        "roster_class": r.get("class", ""),
                        "hold_reason": "학년+이름 중복(양쪽 동수)으로 자동 판정 불가",
                    }
                    transfer_out_hold.append(rec)
                    unresolved_rows.append(rec)
            continue

        # compare에만 있으면 전입 후보
        if compare_count == 1 and roster_count == 0:
            r = compare_group[0]
            base = {"grade": r["grade"], "class": r.get("class", ""), "name": r["name"]}
            compare_only_rows.append(base)
            if r.get("class", ""):
                transfer_in_done.append({**base, "remark": ""})
            else:
                transfer_in_hold.append({**base, "hold_reason": "반 정보 없음"})
            continue

        # roster에만 있으면 전출 후보
        if roster_count == 1 and compare_count == 0:
            r = roster_group[0]
            cls     = r.get("class", "")
            cls_raw = r.get("class_raw", cls)

            if is_excluded_misc_class(cls_raw):
                continue

            base = {"grade": r["grade"], "class": cls, "name": r["name"]}
            roster_only_rows.append(base)
            if cls:
                transfer_out_done.append({**base, "remark": ""})
            else:
                transfer_out_hold.append({**base, "hold_reason": "명부 반 정보 없음"})
            continue

        # 양쪽 1건씩 → 정상 재학생
        if roster_count == 1 and compare_count == 1:
            r = compare_group[0]
            matched_rows.append({
                "grade": r["grade"], "class": r.get("class", ""), "name": r["name"],
            })
            continue

    def _sort_key(r: Dict[str, Any]):
        cls = r.get("class", "")
        try:
            cls_order = (0, int(cls))
        except Exception:
            cls_order = (1, cls)
        return (r.get("grade", 999), cls_order, r.get("name", ""))

    for lst in [matched_rows, compare_only_rows, roster_only_rows, unresolved_rows,
                transfer_in_done, transfer_in_hold, transfer_out_done, transfer_out_hold]:
        lst.sort(key=_sort_key)

    return {
        "matched_rows":      matched_rows,
        "compare_only_rows": compare_only_rows,
        "roster_only_rows":  roster_only_rows,
        "unresolved_rows":   unresolved_rows,
        "transfer_in_done":  transfer_in_done,
        "transfer_in_hold":  transfer_in_hold,
        "transfer_out_done": transfer_out_done,
        "transfer_out_hold": transfer_out_hold,
    }


# =========================
# L4. Scan
# =========================
def scan_diff_pipeline(
    work_root: Path,
    school_name: str,
    target_year: Optional[int],
    school_start_date: date,
    work_date: date,
    roster_basis_date: Optional[date] = None,
    roster_xlsx: Optional[Path] = None,
    col_map: Optional[dict] = None,
    layout_overrides: Optional[dict] = None,
) -> DiffScanResult:

    logs: List[str] = []

    def log(msg: str):
        logs.append(msg)

    target_year = int(target_year) if target_year is not None else int(school_start_date.year)

    sr = DiffScanResult(
        ok=False,
        logs=logs,
        school_name=(school_name or "").strip(),
        year_int=int(target_year),
        year_str=str(target_year),
        school_start_date=school_start_date,
        work_date=work_date,
        roster_basis_date=roster_basis_date,
    )

    try:
        work_root   = Path(work_root).resolve()
        school_name = (school_name or "").strip()

        if not school_name:
            raise ValueError("[ERROR] 학교명을 입력해 주세요.")

        sr.project_root = work_root
        dirs = get_project_dirs(work_root)

        school_dir_candidates = [
            p for p in dirs["SCHOOL_ROOT"].iterdir()
            if p.is_dir() and text_contains(p.name, school_name)
        ]
        if not school_dir_candidates:
            sr.events.append(school_folder_not_found(school_name))
            raise ValueError(f"[ERROR] 학교 폴더를 찾을 수 없습니다. ({school_name})")
        if len(school_dir_candidates) > 1:
            sr.events.append(school_folder_ambiguous(school_name, [p.name for p in school_dir_candidates]))
            raise ValueError(f"[ERROR] 학교 폴더 후보가 여러 개입니다: {[p.name for p in school_dir_candidates]}")

        school_dir = school_dir_candidates[0]
        sr.input_dir  = school_dir
        sr.output_dir = school_dir

        log(f"[OK] 학교 폴더 매칭: {school_dir.name}")

        # 명단 xlsx 기반 학교 존재 확인
        if roster_xlsx:
            _roster_path = Path(roster_xlsx)
            ensure_xlsx_only(_roster_path)
            from core.xlsx_db import school_exists_in_xlsx as _chk
            if not _chk(_roster_path, school_name, col_map):
                sr.events.append(school_not_in_roster(school_name))
                raise ValueError(
                    f"[ERROR] 명단 파일에서 '{school_name}' 학교를 찾을 수 없습니다. "
                    f"(파일: {Path(roster_xlsx).name})"
                )
            sr.roster_xlsx_path = _roster_path
            log(f"[OK] 명단 파일 검증 통과: {_roster_path.name}")
        else:
            log("[WARN] 명단 파일이 지정되지 않았습니다. 학교 존재 확인을 건너뜁니다.")
            sr.roster_xlsx_path = None

        # 템플릿 검증 제거 — diff는 비교 결과 엑셀 직접 출력, 양식 불필요

        # 올해 명부
        roster_wb = None
        try:
            roster_wb, roster_ws, roster_path, roster_year = load_roster_sheet(dirs, school_name)
            sr.roster_path = roster_path
            sr.roster_year = roster_year

            if roster_path is None:
                sr.events.append(roster_not_found())
            raise ValueError("[ERROR] 비교 기준 학생명부를 찾을 수 없습니다.")

            detected_year = parse_roster_year_from_filename(roster_path)
            if detected_year is None:
                detected_year = roster_year

            if detected_year is not None and detected_year != target_year:
                log(
                    f"[WARN] 학생명부 파일명에서 감지한 학년도와 현재 기준 학년도가 다릅니다. "
                    f"(파일명 감지: {detected_year}학년도, 기준: {target_year}학년도)"
                )

            sr.year_int = int(target_year)
            sr.year_str = str(target_year)

            log(f"[OK] 올해 학생명부 감지: {roster_path.name}")

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
                    # roster_date_mismatch는 모달로만 처리 — events 카드에 표시 안 함
                    log(
                        "[DEBUG] 작업일과 명부 기준일이 다릅니다. "
                        f"(작업일={work_date.isoformat()}, 명부 기준일={sr.roster_basis_date.isoformat()})"
                    )

            except Exception as e:
                import traceback
                log(f"[WARN] 학생명부 파일 수정일 조회 중 오류: {e}")
                log(f"[DEBUG] {traceback.format_exc()}")
                sr.roster_basis_date = roster_basis_date or work_date

            basis_date = sr.roster_basis_date or work_date
            if basis_date < school_start_date:
                sr.ref_grade_shift = -1
                log("[INFO] 명부 기준일/개학일 기준으로 '작년 학년도 명부'로 간주합니다.")
            else:
                sr.ref_grade_shift = 0
                log("[INFO] 명부 기준일/개학일 기준으로 '올해 학년도 명부'로 간주합니다.")

        finally:
            if roster_wb is not None:
                roster_wb.close()

        # 비교용 명렬표
        compare_file = find_compare_file(school_dir)
        sr.compare_file = compare_file

        if compare_file is None:
            sr.events.append(compare_file_not_found())
            raise ValueError("[ERROR] 비교용 재학생 명렬표 파일을 찾을 수 없습니다.")

        compare_layout = detect_compare_input_layout(compare_file)
        compare_override = ((layout_overrides or {}).get("compare") or {})
        override_start = compare_override.get("data_start_row")
        if override_start is not None:
            try:
                compare_layout["data_start_row"] = max(1, int(override_start))
                log(f"[INFO] 재학생 명단 시작행을 사용자 입력값으로 적용했습니다. ({compare_layout['data_start_row']}행)")
            except Exception:
                pass
        wb_cmp = safe_load_workbook(compare_file, data_only=True)
        try:
            ws_cmp = wb_cmp.worksheets[0]
            slot_cols = compare_layout.get("slot_cols") or _build_header_slot_map(ws_cmp, compare_layout["header_row"], COMPARE_HEADER_SLOTS)
            issues, issue_rows, row_count, _evts, _marks = validate_compare_input_rows(
                ws_cmp,
                header_row=compare_layout["header_row"],
                data_start_row=compare_layout["data_start_row"],
                slot_cols=slot_cols,
            )
            compare_layout["slot_cols"] = slot_cols
            compare_layout["issue_rows"] = issue_rows
            compare_layout["row_count"] = row_count
            compare_layout["warning"] = issues[0][7:].strip() if issues else ""
            for msg in issues:
                log(msg)
            sr.events.extend(_evts)
            sr.row_marks.extend(_marks)
        finally:
            wb_cmp.close()

        sr.compare_layout = compare_layout
        log(f"[OK] 비교용 재학생 명렬표 감지: {compare_file.name}")
        log(
            f"[DEBUG] 재학생 명렬표 layout: "
            f"header_row={compare_layout['header_row']}, "
            f"data_start_row={compare_layout['data_start_row']}"
        )

        missing_fields: List[str] = []
        if sr.roster_path       is None: missing_fields.append("올해 학생명부")
        if sr.compare_file      is None: missing_fields.append("재학생 명렬표")
        # 템플릿 체크 제거

        sr.missing_fields = missing_fields
        sr.can_execute    = len(missing_fields) == 0
        sr.ok             = True

        log("[DONE] 비교 파이프라인 스캔 완료")
        return sr

    except Exception as e:
        import traceback
        if not isinstance(e, ValueError):
            log(f"[DEBUG] {traceback.format_exc()}")
        log(f"[ERROR] {e}")
        sr.ok         = False
        sr.can_execute = False
        return sr