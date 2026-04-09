"""
core/events.py — 코어 이벤트 타입 정의 + 케이스별 생성 함수

설계 원칙:
  - CoreEvent  : 코어가 판정한 사건 하나. 뱃지·카드의 데이터 원천.
  - RowMark    : 미리보기 테이블 행 색칠 대상. CoreEvent와 code로 연결.
  - 생성 함수  : 케이스 표의 각 행과 1:1 대응. 문구도 여기서 관리.
  - bridge/UI  : 이 객체를 직렬화/렌더링만 함. 문구 조합·logs 파싱 금지.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import List, Literal, Optional

# ──────────────────────────────────────────────
# 타입 정의
# ──────────────────────────────────────────────

EventLevel = Literal["error", "warn", "hold", "info"]

FileKey = Literal[
    "freshmen",
    "transfer_in",
    "transfer_out",
    "teachers",
    "roster",
    "compare",
    "global",
]


@dataclass
class CoreEvent:
    """
    코어가 판정한 사건 하나.
    UI는 이 객체의 message를 그대로 카드에 표시한다.
    logs 파싱이나 문구 재조합 불필요.
    """
    code:       str                    # 이벤트 식별자 — 케이스 표의 code 열과 1:1
    level:      EventLevel             # "error" | "warn" | "hold" | "info"
    message:    str                    # 사용자에게 보여줄 한 줄 요약
    detail:     str        = ""        # 조치 안내 등 부가 설명 (필요할 때만)
    file_key:   FileKey    = "global"  # 어떤 파일의 문제인지
    row:        Optional[int] = None   # 행 번호 — 행 단위 문제일 때만
    field_name: Optional[str] = None   # 열 이름 — 열 단위 문제일 때만
    blocking:   bool       = False     # True면 다음 단계 진행 불가


@dataclass
class RowMark:
    """
    미리보기 테이블에서 행 색칠 대상.
    왜 칠하는지(code)가 포함되어 있어 카드와 같은 원천을 공유한다.
    """
    file_key: FileKey
    row:      int           # 엑셀 기준 절대 행 번호
    level:    EventLevel    # "error"=빨강, "warn"=노랑, "hold"=분홍
    code:     str           # 연결된 CoreEvent.code


# ──────────────────────────────────────────────
# 설정 / 시작 전
# ──────────────────────────────────────────────

def resources_config_error(detail: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "RESOURCES_CONFIG_ERROR",
        level    = "error",
        message  = "resources 폴더에 필요한 자료가 모두 포함되었는지 확인하세요.",
        detail   = detail,
        blocking = True,
    )


def roster_xls_format() -> CoreEvent:
    return CoreEvent(
        code     = "ROSTER_XLS_FORMAT",
        level    = "error",
        message  = "명단 파일은 .xlsx 형식이어야 합니다. Excel에서 .xlsx로 저장 후 다시 선택해 주세요.",
        blocking = True,
    )


def school_not_in_roster(school_name: str) -> CoreEvent:
    return CoreEvent(
        code     = "SCHOOL_NOT_IN_ROSTER",
        level    = "error",
        message  = f"명단 파일에서 '{school_name}' 학교를 찾을 수 없습니다.",
        blocking = True,
    )


# ──────────────────────────────────────────────
# 학교 / 폴더 문제
# ──────────────────────────────────────────────

def school_folder_not_found(school_name: str) -> CoreEvent:
    return CoreEvent(
        code     = "SCHOOL_FOLDER_NOT_FOUND",
        level    = "error",
        message  = f"'{school_name}' 학교 폴더를 찾을 수 없습니다.",
        blocking = True,
    )


def school_folder_ambiguous(school_name: str, candidates: list) -> CoreEvent:
    names = ", ".join(str(c) for c in candidates)
    return CoreEvent(
        code     = "SCHOOL_FOLDER_AMBIGUOUS",
        level    = "error",
        message  = f"'{school_name}' 학교 폴더 후보가 여러 개입니다: {names}",
        blocking = True,
    )


def no_input_files() -> CoreEvent:
    return CoreEvent(
        code     = "NO_INPUT_FILES",
        level    = "error",
        message  = "학교 폴더 내 명단 파일을 하나도 찾을 수 없습니다.",
        detail   = "신입생, 전입생, 전출생, 교직원 파일 중 하나 이상 필요합니다.",
        blocking = True,
    )


def input_xls_format(file_names: list) -> CoreEvent:
    names = ", ".join(file_names)
    return CoreEvent(
        code     = "INPUT_XLS_FORMAT",
        level    = "error",
        message  = f".xls 형식은 지원하지 않습니다. .xlsx로 저장 후 다시 시도해 주세요.",
        detail   = f"해당 파일: {names}",
        blocking = True,
    )


def duplicate_input_file(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "DUPLICATE_INPUT_FILE",
        level    = "error",
        message  = f"{kind_label} 명단 파일이 2개 이상 감지되었습니다. 하나만 남겨 주세요.",
        file_key = file_key,
        blocking = True,
    )


# ──────────────────────────────────────────────
# 파일 구조 오류
# ──────────────────────────────────────────────

def missing_header(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "MISSING_HEADER",
        level    = "error",
        message  = f"{kind_label} 파일에서 헤더를 찾을 수 없습니다.",
        file_key = file_key,
        blocking = False,
    )


def missing_data_start(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "MISSING_DATA_START",
        level    = "error",
        message  = f"{kind_label} 데이터 시작 행을 찾을 수 없습니다.",
        detail   = "헤더 행 아래에 실제 데이터가 있는지 확인해 주세요.",
        file_key = file_key,
        blocking = False,
    )


def empty_data(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "EMPTY_DATA",
        level    = "error",
        message  = f"{kind_label} 파일에서 데이터 행을 찾을 수 없습니다.",
        file_key = file_key,
        blocking = False,
    )


def missing_required_col(file_key: FileKey, kind_label: str, col_name: str) -> CoreEvent:
    return CoreEvent(
        code       = "MISSING_REQUIRED_COL",
        level      = "error",
        message    = f"{kind_label} 파일에서 '{col_name}' 열을 찾을 수 없습니다.",
        file_key   = file_key,
        field_name = col_name,
        blocking   = True,
    )


# ──────────────────────────────────────────────
# 파일 내용 경고 + RowMark
# ──────────────────────────────────────────────

def grade_format_warn(
    file_key: FileKey, kind_label: str, row: int, reason: str
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code       = "GRADE_FORMAT_WARN",
        level      = "warn",
        message    = f"{kind_label} {row}행 '학년' 열 — {reason}",
        file_key   = file_key,
        row        = row,
        field_name = "학년",
    )
    mark = RowMark(file_key=file_key, row=row, level="warn", code="GRADE_FORMAT_WARN")
    return event, mark


def class_format_warn(
    file_key: FileKey, kind_label: str, row: int, reason: str
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code       = "CLASS_FORMAT_WARN",
        level      = "warn",
        message    = f"{kind_label} {row}행 '반' 열 — {reason}",
        file_key   = file_key,
        row        = row,
        field_name = "반",
    )
    mark = RowMark(file_key=file_key, row=row, level="warn", code="CLASS_FORMAT_WARN")
    return event, mark


def name_format_warn(
    file_key: FileKey, kind_label: str, row: int, reason: str
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code       = "NAME_FORMAT_WARN",
        level      = "warn",
        message    = f"{kind_label} {row}행 '이름' 열 — {reason}",
        file_key   = file_key,
        row        = row,
        field_name = "이름",
    )
    mark = RowMark(file_key=file_key, row=row, level="warn", code="NAME_FORMAT_WARN")
    return event, mark


def empty_required_field(
    file_key: FileKey, kind_label: str, row: int, field_name: str
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code       = "EMPTY_REQUIRED_FIELD",
        level      = "warn",
        message    = f"{kind_label} {row}행 '{field_name}' 값이 비어 있습니다.",
        file_key   = file_key,
        row        = row,
        field_name = field_name,
    )
    mark = RowMark(file_key=file_key, row=row, level="warn", code="EMPTY_REQUIRED_FIELD")
    return event, mark


def empty_row(
    file_key: FileKey, kind_label: str, row: int
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code     = "EMPTY_ROW",
        level    = "warn",
        message  = f"{kind_label} {row}행이 비어 있습니다.",
        file_key = file_key,
        row      = row,
    )
    mark = RowMark(file_key=file_key, row=row, level="warn", code="EMPTY_ROW")
    return event, mark


def merged_cell(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "MERGED_CELL",
        level    = "warn",
        message  = f"{kind_label} 데이터 영역에 병합된 셀이 있습니다.",
        detail   = "병합된 셀은 읽기 오류를 유발할 수 있습니다. 병합을 해제 후 다시 시도해 주세요.",
        file_key = file_key,
    )


def multiple_sheets(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "MULTIPLE_SHEETS",
        level    = "warn",
        message  = f"{kind_label} 시트가 두 개 이상입니다. 첫 번째 시트만 사용합니다.",
        file_key = file_key,
    )


def kindergarten_in_file(file_key: FileKey, kind_label: str) -> CoreEvent:
    return CoreEvent(
        code     = "KINDERGARTEN_IN_FILE",
        level    = "warn",
        message  = f"{kind_label} 파일에 유치부 학생이 포함되어 있습니다.",
        file_key = file_key,
    )


def no_teacher_id_request() -> CoreEvent:
    return CoreEvent(
        code     = "NO_TEACHER_ID_REQUEST",
        level    = "warn",
        message  = "교직원 명단에 관리용 아이디 신청자가 한 건도 없습니다. 관리용 ID 신청 열을 확인해 주세요.",
        file_key = "teachers",
    )


# ──────────────────────────────────────────────
# 학교 구분 / 명부
# ──────────────────────────────────────────────

def school_kind_unknown() -> CoreEvent:
    return CoreEvent(
        code     = "SCHOOL_KIND_UNKNOWN",
        level    = "warn",
        message  = "자동으로 학교를 판정할 수 없습니다. 학교 구분을 직접 선택해 주세요.",
        blocking = False,
    )


def roster_not_found(reason: str = "") -> CoreEvent:
    """reason: '전입' | '전출' | '전입/전출' | '신입생 타학년' | '' """
    if reason in ("전입", "전출", "전입/전출"):
        message = f"{reason}생 처리를 위해 학생 명부가 필요합니다. 학교 폴더에 명부를 추가해 주세요."
    elif reason == "신입생 타학년":
        message = "신입생 파일에 1학년 외 학년이 포함되어 있어 명부가 필요합니다. 학교 폴더에 명부를 추가해 주세요."
    else:
        message = "작업을 위해 학생 명부가 필요합니다. 학교 폴더에 명부를 추가해 주세요."
    return CoreEvent(
        code     = "ROSTER_NOT_FOUND",
        level    = "error",
        message  = message,
        file_key = "roster",
        blocking = False,
    )


def freshmen_extra_grades_info(grades: list) -> CoreEvent:
    """신입생 파일에 1학년 외 학년 포함 — WARN 이벤트"""
    grade_str = ", ".join(f"{g}학년" for g in grades)
    return CoreEvent(
        code     = "FRESHMEN_EXTRA_GRADES",
        level    = "warn",
        message  = f"신입생 파일에 {grade_str}이(가) 포함되어 있습니다.",
        file_key = "freshmen",
    )


def freshmen_no_roster_manual() -> CoreEvent:
    """신입생 타학년 + 명부 없음 → 학년도 직접 입력 안내 (hold)"""
    return CoreEvent(
        code     = "FRESHMEN_NO_ROSTER_MANUAL",
        level    = "hold",
        message  = "학교 폴더에 명부를 추가하거나, 사이드바에서 학년도 아이디 규칙을 직접 입력하세요.",
        file_key = "freshmen",
    )


def roster_not_found_at_run() -> CoreEvent:
    """실행 중 명부를 찾을 수 없는 경우 — 스캔 후 파일이 삭제된 상황"""
    return CoreEvent(
        code     = "ROSTER_NOT_FOUND_AT_RUN",
        level    = "error",
        message  = "학생 명부를 찾을 수 없습니다.",
        detail   = "스캔 후 명부 파일이 삭제되었을 수 있습니다. 다시 스캔해 주세요.",
        file_key = "roster",
        blocking = True,
    )


def roster_date_mismatch(basis_date: str, work_date: str) -> CoreEvent:
    return CoreEvent(
        code     = "ROSTER_DATE_MISMATCH",
        level    = "warn",
        message  = "학생명부 기준일과 작업일이 다릅니다. 어느 날짜를 기준일로 사용할까요?",
        detail   = f"명부 기준일: {basis_date} / 작업일: {work_date}",
        file_key = "roster",
    )


def open_date_missing() -> CoreEvent:
    return CoreEvent(
        code     = "OPEN_DATE_MISSING",
        level    = "warn",
        message  = "전출 파일이 있습니다. 실행 전 개학일을 입력해 주세요.",
        blocking = False,
    )


def _class_token(raw: str | int | None, grade: int | None = None) -> str:
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    s = s.replace("　", " ").replace(" ", " ")
    s = re.sub(r"\s+", "", s)

    if s in {"유치원", "유치원반"}:
        return "유치원"

    s = re.sub(r"반+$", "", s)

    m = re.fullmatch(r"(\d+)-(\d+)", s)
    if m:
        return f"{int(m.group(1))}-{int(m.group(2))}"

    m = re.fullmatch(r"(\d+)", s)
    if m:
        cls = int(m.group(1))
        if grade:
            return f"{int(grade)}-{cls}"
        return str(cls)

    return s


def format_school_info(grade: int | None = None, class_: str | int | None = None) -> str:
    token = _class_token(class_, grade)
    if token:
        return token
    if grade:
        return str(int(grade))
    return ""


TRANSFER_IN_REASON_TEXT = {
    "PREFIX_MODE_UNAVAILABLE": "ID 규칙을 확인할 수 없어 확인이 필요합니다.",
    "DUP_WITH_FRESHMEN": "신입생 명단과 학년/반/이름이 동일합니다.",
    "ROSTER_DUPLICATE": "학생명부에 이미 존재하는 학생입니다.",
    "ROSTER_SUSPECT_SAME_PERSON": "학생명부에 동일인으로 의심되는 학생이 있습니다.",
}

TRANSFER_OUT_REASON_TEXT = {
    "NAME_KEY_EMPTY": "이름을 확인할 수 없어 확인이 필요합니다.",
    "ROSTER_DUPLICATE_NAME": "학생명부에 동명이인이 있어 확인이 필요합니다.",
    "GRADE_NAME_MULTI_MATCH": "학생명부에서 후보가 여러 건 확인되어 확인이 필요합니다.",
    "ROSTER_NOT_FOUND": "학생명부에서 확인되지 않아 확인이 필요합니다.",
    "MULTIPLE_MATCHES": "학생명부에서 중복 매칭되어 확인이 필요합니다.",
    "AUTO_SKIP_NOT_FOUND": "학생명부에 존재하지 않아 자동 제외되었습니다.",
}

DIFF_REASON_TEXT = {
    "COMPARE_DUPLICATE_NAME": "재학생 명단에 동명이인이 있어 확인이 필요합니다.",
    "ROSTER_DUPLICATE_NAME": "명부에 동명이인이 있어 확인이 필요합니다.",
    "BOTH_DUPLICATE_NAME": "재학생 명단과 명부 양쪽에 동명이인이 있어 확인이 필요합니다.",
    "ROSTER_SUFFIX_DUPLICATES": "명부에 동명이인(A,B,C 등) 후보가 있어 확인이 필요합니다.",
    "ROSTER_SUSPECT_SAME_PERSON": "명부에 동일인으로 의심되는 후보가 있어 확인이 필요합니다.",
    "MISSING_CLASS": "반 정보가 없어 확인이 필요합니다.",
    "MISSING_ROSTER_CLASS": "명부 반 정보가 없어 확인이 필요합니다.",
}


def transfer_in_reason_text(code: str) -> str:
    return TRANSFER_IN_REASON_TEXT.get(code, code)


def transfer_out_reason_text(code: str) -> str:
    return TRANSFER_OUT_REASON_TEXT.get(code, code)


def diff_reason_text(code: str) -> str:
    return DIFF_REASON_TEXT.get(code, code)


def _student_event_message(prefix: str, name: str, reason: str, grade: int | None = None, class_: str | int | None = None) -> str:
    info = format_school_info(grade, class_)
    if info:
        return f"{prefix}: {info} {name} - {reason}"
    return f"{prefix}: {name} - {reason}"


def freshmen_transfer_same_student_scan(grade: int, class_: str, name: str) -> CoreEvent:
    return CoreEvent(
        code     = "FRESHMEN_TRANSFER_SAME_STUDENT_SCAN",
        level    = "warn",
        message  = _student_event_message(
            "중복 감지",
            name,
            "신입생 명단과 전입생 명단에 동일한 학적정보가 있습니다.",
            grade,
            class_,
        ),
        file_key = "global",
        blocking = False,
    )


# ──────────────────────────────────────────────
# 실행 완료 — hold
# ──────────────────────────────────────────────

def transfer_in_hold(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "TRANSFER_IN_HOLD",
        level    = "hold",
        message  = _student_event_message("전입생 보류", name, transfer_in_reason_text(reason_code), grade, class_),
        file_key = "transfer_in",
    )


def transfer_out_hold(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "TRANSFER_OUT_HOLD",
        level    = "hold",
        message  = _student_event_message("전출생 보류", name, transfer_out_reason_text(reason_code), grade, class_),
        file_key = "transfer_out",
    )


def freshmen_transfer_dup(name: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "FRESHMEN_TRANSFER_DUP",
        level    = "hold",
        message  = _student_event_message(
            "전입생 보류",
            name,
            "신입생 명단과 학년/반/이름이 동일합니다.",
            grade,
            class_,
        ),
        file_key = "transfer_in",
    )


def roster_duplicate_transfer(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    if reason_code == "ROSTER_SUSPECT_SAME_PERSON":
        detail = "명부에 같은 학년·이름의 학생이 있으나 반이 다릅니다. 동일인인지 확인 후 처리해 주세요."
    else:
        detail = "이미 명부에 존재하는 학생입니다. 확인 후 처리해 주세요."
    return CoreEvent(
        code     = "ROSTER_DUPLICATE_TRANSFER",
        level    = "hold",
        message  = _student_event_message("전입생 보류", name, transfer_in_reason_text(reason_code), grade, class_),
        detail   = detail,
        file_key = "transfer_in",
    )


def duplicate_name(count: int) -> CoreEvent:
    return CoreEvent(
        code    = "DUPLICATE_NAME",
        level   = "warn",
        message = f"동명이인 경고: {count}건 - 실행 결과 탭 필터에서 확인해 주세요.",
    )


# ──────────────────────────────────────────────
# 실행 중 오류
# ──────────────────────────────────────────────

def template_register_not_found() -> CoreEvent:
    return CoreEvent(
        code     = "TEMPLATE_REGISTER_NOT_FOUND",
        level    = "error",
        message  = "등록 템플릿 파일을 찾을 수 없습니다.",
        blocking = True,
    )


def template_notice_not_found() -> CoreEvent:
    return CoreEvent(
        code     = "TEMPLATE_NOTICE_NOT_FOUND",
        level    = "error",
        message  = "안내 템플릿 파일을 찾을 수 없습니다.",
        blocking = True,
    )


def db_file_error(detail: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "DB_FILE_ERROR",
        level    = "error",
        message  = "DB 폴더에 '학교전체명단' xlsb 파일 문제.",
        detail   = detail,
        blocking = True,
    )


def open_date_required() -> CoreEvent:
    return CoreEvent(
        code     = "OPEN_DATE_REQUIRED",
        level    = "error",
        message  = "전출 처리에 필요한 개학일이 입력되지 않았습니다.",
        blocking = True,
    )


# ──────────────────────────────────────────────
# diff 전용
# ──────────────────────────────────────────────

def compare_file_not_found() -> CoreEvent:
    return CoreEvent(
        code     = "COMPARE_FILE_NOT_FOUND",
        level    = "error",
        message  = "비교용 재학생 명렬표 파일을 찾을 수 없습니다.",
        file_key = "compare",
        blocking = True,
    )


def compare_file_format_warn(
    row: int, field_name: str, reason: str
) -> tuple[CoreEvent, RowMark]:
    event = CoreEvent(
        code       = "COMPARE_FORMAT_WARN",
        level      = "warn",
        message    = f"재학생 파일 {row}행 '{field_name}' — {reason}",
        file_key   = "compare",
        row        = row,
        field_name = field_name,
    )
    mark = RowMark(file_key="compare", row=row, level="warn", code="COMPARE_FORMAT_WARN")
    return event, mark


def diff_unresolved(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "DIFF_UNRESOLVED",
        level    = "hold",
        message  = _student_event_message("자동 판정 불가", name, diff_reason_text(reason_code), grade, class_),
        file_key = "compare",
    )


def diff_transfer_in_hold(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "DIFF_TRANSFER_IN_HOLD",
        level    = "hold",
        message  = _student_event_message("전입생 보류", name, diff_reason_text(reason_code), grade, class_),
        file_key = "transfer_in",
    )


def diff_transfer_out_hold(name: str, reason_code: str, grade: int = 0, class_: str = "") -> CoreEvent:
    return CoreEvent(
        code     = "DIFF_TRANSFER_OUT_HOLD",
        level    = "hold",
        message  = _student_event_message("전출생 보류", name, diff_reason_text(reason_code), grade, class_),
        file_key = "transfer_out",
    )


# ──────────────────────────────────────────────
# 헬퍼 — Result 객체에서 status 계산
# ──────────────────────────────────────────────

def compute_status(events: List[CoreEvent]) -> str:
    """events 리스트에서 전체 상태를 계산. bridge/UI에서 사용."""
    if any(e.level == "error" for e in events):
        return "error"
    if any(e.level == "hold"  for e in events):
        return "hold"
    if any(e.level == "warn"  for e in events):
        return "warn"
    return "ok"
