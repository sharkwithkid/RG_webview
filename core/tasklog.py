from __future__ import annotations

import csv
import json
import time
from dataclasses import dataclass, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import io


TASKLOG_HEADERS = [
    "seq",
    "logged_at",
    "pipeline_kind",          # main | diff

    "school_name",
    "school_folder_name",

    "worker",
    "mail_received_date",
    "work_date",
    "open_date",
    "roster_basis_date",

    "status",                 # 완료 | 보류있음 | 실패

    "freshmen_status",
    "transfer_in_status",
    "transfer_out_status",
    "teacher_status",

    "transfer_in_done",
    "transfer_in_hold",
    "transfer_out_done",
    "transfer_out_hold",
    "transfer_out_auto_skip",

    # 기본 진행 단계
    "file_created",
    "file_created_at",
    "bm_registered",
    "bm_registered_at",
    "email_sent",
    "email_sent_at",
    "sms_sent",
    "sms_sent_at",

    # 추가 작업 축
    "extra_work_needed",
    "extra_work_done",
    "extra_work_done_at",
    "extra_work_note",

    "output_files",
    "note",
    "error_summary",
]


# 내보내기용
FORMAT_LOG_HEADERS = [
    "이메일 도착일자",
    "완료 이메일발송",
    "작업자",
    "작업현황(신입생)",
    "작업현황(전입생)",
    "작업현황(전출생)",
    "작업현황(교직원)",
    "자료실 순번",
]


# =========================================================
# Public API for UI layer
# =========================================================
# - create_main_tasklog
# - create_diff_tasklog
# - update_tasklog_progress_by_seq
# - find_tasklog_row_by_seq
# - list_school_progress_rows
#
# 아래의 나머지 함수들은 tasklog 내부 동작용 helper로 본다.
# UI(PyQt / Streamlit)에서는 직접 호출하지 않는다.
# =========================================================

@dataclass
class TaskLogEntry:
    seq: int
    logged_at: str
    pipeline_kind: str

    school_name: str
    school_folder_name: str

    worker: str
    mail_received_date: str
    work_date: str
    open_date: str
    roster_basis_date: str

    status: str

    freshmen_status: str
    transfer_in_status: str
    transfer_out_status: str
    teacher_status: str

    transfer_in_done: int
    transfer_in_hold: int
    transfer_out_done: int
    transfer_out_hold: int
    transfer_out_auto_skip: int

    file_created: str
    file_created_at: str
    bm_registered: str
    bm_registered_at: str
    email_sent: str
    email_sent_at: str
    sms_sent: str
    sms_sent_at: str

    extra_work_needed: str
    extra_work_done: str
    extra_work_done_at: str
    extra_work_note: str

    output_files: str
    note: str
    error_summary: str


def _date_to_str(d: Optional[date]) -> str:
    return "" if d is None else d.isoformat()


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _yn(flag: Optional[bool]) -> str:
    if flag is None:
        return ""
    return "Y" if flag else "N"


def get_tasklog_dir(work_root: Path) -> Path:
    return Path(work_root).resolve() / "_tasklog"


def get_tasklog_csv_path(work_root: Path) -> Path:
    return get_tasklog_dir(work_root) / "tasklog.csv"


def get_tasklog_lock_path(work_root: Path) -> Path:
    return get_tasklog_dir(work_root) / "tasklog.lock"


def ensure_tasklog_store(work_root: Path) -> Path:
    log_dir = get_tasklog_dir(work_root)
    log_dir.mkdir(parents=True, exist_ok=True)

    csv_path = get_tasklog_csv_path(work_root)
    if not csv_path.exists():
        with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=TASKLOG_HEADERS)
            writer.writeheader()
    return csv_path


def _acquire_lock(lock_path: Path, timeout_sec: float = 5.0, interval_sec: float = 0.1) -> None:
    start = time.time()
    while lock_path.exists():
        if time.time() - start > timeout_sec:
            raise TimeoutError("[ERROR] tasklog.lock 획득에 실패했습니다.")
        time.sleep(interval_sec)
    lock_path.write_text(str(datetime.now()), encoding="utf-8")


def _release_lock(lock_path: Path) -> None:
    if lock_path.exists():
        lock_path.unlink()


def _read_csv_rows_with_fallback(csv_path: Path) -> List[Dict[str, str]]:
    raw = csv_path.read_bytes()

    text = None
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue

    if text is None:
        text = raw.decode("utf-8", errors="replace")

    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


def read_tasklog_rows(work_root: Path) -> List[Dict[str, str]]:
    csv_path = ensure_tasklog_store(work_root)
    return _read_csv_rows_with_fallback(csv_path)


def write_tasklog_rows(work_root: Path, rows: List[Dict[str, str]]) -> None:
    csv_path = ensure_tasklog_store(work_root)
    lock_path = get_tasklog_lock_path(work_root)

    _acquire_lock(lock_path)
    try:
        tmp_path = csv_path.with_suffix(".tmp")

        with tmp_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=TASKLOG_HEADERS)
            writer.writeheader()
            for row in rows:
                safe_row = {k: row.get(k, "") for k in TASKLOG_HEADERS}
                writer.writerow(safe_row)

        tmp_path.replace(csv_path)
    finally:
        _release_lock(lock_path)


def append_new_tasklog_entry(
    work_root: Path,
    *,
    entry_builder,
) -> TaskLogEntry:
    csv_path = ensure_tasklog_store(work_root)
    lock_path = get_tasklog_lock_path(work_root)

    _acquire_lock(lock_path)
    try:
        rows = _read_csv_rows_with_fallback(csv_path)
        max_seq = 0
        for row in rows:
            try:
                max_seq = max(max_seq, int(row.get("seq", 0)))
            except Exception:
                continue

        seq = max_seq + 1
        entry = entry_builder(seq)

        with csv_path.open("a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=TASKLOG_HEADERS)
            writer.writerow(asdict(entry))

        return entry
    finally:
        _release_lock(lock_path)


def find_latest_school_log(work_root: Path, school_name: str) -> Optional[Dict[str, str]]:
    school_name = (school_name or "").strip()
    if not school_name:
        return None

    rows = read_tasklog_rows(work_root)
    matched = [r for r in rows if (r.get("school_name") or "").strip() == school_name]
    if not matched:
        return None

    def _seq_key(row: Dict[str, str]) -> int:
        try:
            return int(row.get("seq", 0))
        except Exception:
            return 0

    matched.sort(key=_seq_key, reverse=True)
    return matched[0]


def find_tasklog_row_by_seq(work_root: Path, seq: int) -> Optional[Dict[str, str]]:
    rows = read_tasklog_rows(work_root)
    for row in rows:
        try:
            if int(row.get("seq", 0)) == int(seq):
                return row
        except Exception:
            continue
    return None


def update_tasklog_progress_by_seq(
    work_root: Path,
    seq: int,
    *,
    bm_registered: Optional[bool] = None,
    email_sent: Optional[bool] = None,
    sms_sent: Optional[bool] = None,
    extra_work_needed: Optional[bool] = None,
    extra_work_done: Optional[bool] = None,
    extra_work_note: Optional[str] = None,
) -> bool:
    rows = read_tasklog_rows(work_root)
    found = False

    for row in rows:
        try:
            if int(row.get("seq", 0)) != int(seq):
                continue
        except Exception:
            continue

        found = True

        if bm_registered is not None:
            row["bm_registered"] = _yn(bm_registered)
            row["bm_registered_at"] = _now_str() if bm_registered else ""

        if email_sent is not None:
            row["email_sent"] = _yn(email_sent)
            row["email_sent_at"] = _now_str() if email_sent else ""

        if sms_sent is not None:
            row["sms_sent"] = _yn(sms_sent)
            row["sms_sent_at"] = _now_str() if sms_sent else ""

        if extra_work_needed is not None:
            row["extra_work_needed"] = _yn(extra_work_needed)
            if not extra_work_needed:
                row["extra_work_done"] = "N"
                row["extra_work_done_at"] = ""
                row["extra_work_note"] = ""

        if extra_work_done is not None:
            row["extra_work_done"] = _yn(extra_work_done)
            row["extra_work_done_at"] = _now_str() if extra_work_done else ""

        if extra_work_note is not None:
            row["extra_work_note"] = (extra_work_note or "").strip()

        break

    if found:
        write_tasklog_rows(work_root, rows)

    return found


def get_format_log_csv_path(work_root: Path) -> Path:
    return get_tasklog_dir(work_root) / "tasklog_export_format.csv"


def _is_done_status(value: Optional[str]) -> bool:
    return (value or "").strip() == "완료"


def _done_or_blank(flag: bool) -> str:
    return "완료" if flag else ""


def convert_tasklog_row_to_format_row(row: Dict[str, str]) -> Dict[str, str]:
    """
    시스템 로그 1행 -> 공유폴더 양식 로그 1행 변환

    규칙
    - 이메일 도착일자 -> mail_received_date
    - 완료 이메일발송 -> email_sent == Y 이면 '완료'
    - 작업자 -> worker
    - 작업현황(신입생/전입생/전출생/교직원) -> 각 status가 '완료'면 '완료', 아니면 공란
    - 자료실 순번 -> 보류이므로 공란
    - 추가작업은 자동 채움 대상에서 제외
    """
    return {
        "이메일 도착일자": (row.get("mail_received_date") or "").strip(),
        "완료 이메일발송": _done_or_blank((row.get("email_sent") or "").strip() == "Y"),
        "작업자": (row.get("worker") or "").strip(),
        "작업현황(신입생)": _done_or_blank(_is_done_status(row.get("freshmen_status"))),
        "작업현황(전입생)": _done_or_blank(_is_done_status(row.get("transfer_in_status"))),
        "작업현황(전출생)": _done_or_blank(_is_done_status(row.get("transfer_out_status"))),
        "작업현황(교직원)": _done_or_blank(_is_done_status(row.get("teacher_status"))),
        "자료실 순번": "",
    }


def build_format_log_rows(
    work_root: Path,
    *,
    include_failed: bool = False,
) -> List[Dict[str, str]]:
    rows = read_tasklog_rows(work_root)

    if not include_failed:
        rows = [r for r in rows if (r.get("status") or "").strip() != "실패"]

    def _seq_key(row: Dict[str, str]) -> int:
        try:
            return int(row.get("seq", 0))
        except Exception:
            return 0

    rows.sort(key=_seq_key)
    return [convert_tasklog_row_to_format_row(r) for r in rows]


def export_format_log_csv(
    work_root: Path,
    *,
    include_failed: bool = False,
) -> Path:
    out_path = get_format_log_csv_path(work_root)
    rows = build_format_log_rows(work_root, include_failed=include_failed)
    lock_path = get_tasklog_lock_path(work_root)

    _acquire_lock(lock_path)
    try:
        with out_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=FORMAT_LOG_HEADERS)
            writer.writeheader()
            for row in rows:
                safe_row = {k: row.get(k, "") for k in FORMAT_LOG_HEADERS}
                writer.writerow(safe_row)
    finally:
        _release_lock(lock_path)

    return out_path


def build_main_tasklog_entry(
    *,
    seq: int,
    school_name: str,
    school_folder_name: str,
    worker: str,
    mail_received_date: Optional[date],
    work_date: Optional[date],
    open_date: Optional[date],
    roster_basis_date: Optional[date],
    scan: Any,
    result: Any,
    note: str = "",
) -> TaskLogEntry:
    ti_done = int(getattr(result, "transfer_in_done", 0))
    ti_hold = int(getattr(result, "transfer_in_hold", 0))
    to_done = int(getattr(result, "transfer_out_done", 0))
    to_hold = int(getattr(result, "transfer_out_hold", 0))
    to_auto_skip = int(getattr(result, "transfer_out_auto_skip", 0))

    if not getattr(result, "ok", False):
        status = "실패"
    elif ti_hold > 0 or max(to_hold - to_auto_skip, 0) > 0:
        status = "보류있음"
    else:
        status = "완료"

    freshmen_file = getattr(scan, "freshmen_file", None) if scan is not None else None
    teacher_file = getattr(scan, "teacher_file", None) if scan is not None else None
    transfer_file = getattr(scan, "transfer_file", None) if scan is not None else None
    withdraw_file = getattr(scan, "withdraw_file", None) if scan is not None else None

    if freshmen_file:
        freshmen_status = "완료" if getattr(result, "ok", False) else "실패"
    else:
        freshmen_status = ""

    if teacher_file:
        teacher_status = "완료" if getattr(result, "ok", False) else "실패"
    else:
        teacher_status = ""

    if transfer_file:
        if not getattr(result, "ok", False):
            transfer_in_status = "실패"
        elif ti_hold > 0:
            transfer_in_status = "보류"
        else:
            transfer_in_status = "완료"
    else:
        transfer_in_status = ""

    if withdraw_file:
        real_out_hold = max(to_hold - to_auto_skip, 0)
        if not getattr(result, "ok", False):
            transfer_out_status = "실패"
        elif real_out_hold > 0:
            transfer_out_status = "보류"
        else:
            transfer_out_status = "완료"
    else:
        transfer_out_status = ""

    outputs = [str(p) for p in getattr(result, "outputs", [])]

    error_summary = ""
    if not getattr(result, "ok", False):
        logs = getattr(result, "logs", []) or []
        error_lines = [x for x in logs if "[ERROR]" in str(x)]
        if error_lines:
            error_summary = "\n".join(error_lines[-5:])
        else:
            error_summary = "\n".join([str(x) for x in logs[-5:]])

    file_created = "Y" if getattr(result, "ok", False) else "N"
    file_created_at = _now_str() if getattr(result, "ok", False) else ""

    return TaskLogEntry(
        seq=seq,
        logged_at=_now_str(),
        pipeline_kind="main",

        school_name=(school_name or "").strip(),
        school_folder_name=(school_folder_name or "").strip(),

        worker=(worker or "").strip(),
        mail_received_date=_date_to_str(mail_received_date),
        work_date=_date_to_str(work_date),
        open_date=_date_to_str(open_date),
        roster_basis_date=_date_to_str(roster_basis_date),

        status=status,

        freshmen_status=freshmen_status,
        transfer_in_status=transfer_in_status,
        transfer_out_status=transfer_out_status,
        teacher_status=teacher_status,

        transfer_in_done=ti_done,
        transfer_in_hold=ti_hold,
        transfer_out_done=to_done,
        transfer_out_hold=to_hold,
        transfer_out_auto_skip=to_auto_skip,

        file_created=file_created,
        file_created_at=file_created_at,
        bm_registered="N",
        bm_registered_at="",
        email_sent="N",
        email_sent_at="",
        sms_sent="N",
        sms_sent_at="",

        extra_work_needed="N",
        extra_work_done="N",
        extra_work_done_at="",
        extra_work_note="",

        output_files=json.dumps(outputs, ensure_ascii=False),
        note=(note or "").strip(),
        error_summary=error_summary,
    )


def build_diff_tasklog_entry(
    *,
    seq: int,
    school_name: str,
    school_folder_name: str,
    worker: str,
    mail_received_date: Optional[date],
    work_date: Optional[date],
    open_date: Optional[date],
    result: Any,
    note: str = "",
) -> TaskLogEntry:
    ti_done = int(getattr(result, "transfer_in_done", 0))
    ti_hold = int(getattr(result, "transfer_in_hold", 0))
    to_done = int(getattr(result, "transfer_out_done", 0))
    to_hold = int(getattr(result, "transfer_out_hold", 0))

    if not getattr(result, "ok", False):
        status = "실패"
    elif ti_hold > 0 or to_hold > 0:
        status = "보류있음"
    else:
        status = "완료"

    if not getattr(result, "ok", False):
        transfer_in_status = "실패"
        transfer_out_status = "실패"
    else:
        transfer_in_status = "보류" if ti_hold > 0 else "완료"
        transfer_out_status = "보류" if to_hold > 0 else "완료"

    outputs = [str(p) for p in getattr(result, "outputs", [])]

    error_summary = ""
    if not getattr(result, "ok", False):
        logs = getattr(result, "logs", []) or []
        error_lines = [x for x in logs if "[ERROR]" in str(x)]
        if error_lines:
            error_summary = "\n".join(error_lines[-5:])
        else:
            error_summary = "\n".join([str(x) for x in logs[-5:]])

    file_created = "Y" if getattr(result, "ok", False) else "N"
    file_created_at = _now_str() if getattr(result, "ok", False) else ""

    return TaskLogEntry(
        seq=seq,
        logged_at=_now_str(),
        pipeline_kind="diff",

        school_name=(school_name or "").strip(),
        school_folder_name=(school_folder_name or "").strip(),

        worker=(worker or "").strip(),
        mail_received_date=_date_to_str(mail_received_date),
        work_date=_date_to_str(work_date),
        open_date=_date_to_str(open_date),
        roster_basis_date="",

        status=status,

        freshmen_status="",
        transfer_in_status=transfer_in_status,
        transfer_out_status=transfer_out_status,
        teacher_status="",

        transfer_in_done=ti_done,
        transfer_in_hold=ti_hold,
        transfer_out_done=to_done,
        transfer_out_hold=to_hold,
        transfer_out_auto_skip=0,

        file_created=file_created,
        file_created_at=file_created_at,
        bm_registered="N",
        bm_registered_at="",
        email_sent="N",
        email_sent_at="",
        sms_sent="N",
        sms_sent_at="",

        extra_work_needed="N",
        extra_work_done="N",
        extra_work_done_at="",
        extra_work_note="",

        output_files=json.dumps(outputs, ensure_ascii=False),
        note=(note or "").strip(),
        error_summary=error_summary,
    )


def create_main_tasklog(
    work_root: Path,
    *,
    school_name: str,
    school_folder_name: str,
    worker: str,
    mail_received_date: Optional[date],
    work_date: Optional[date],
    open_date: Optional[date],
    roster_basis_date: Optional[date],
    scan: Any,
    result: Any,
    note: str = "",
) -> TaskLogEntry:
    return append_new_tasklog_entry(
        work_root,
        entry_builder=lambda seq: build_main_tasklog_entry(
            seq=seq,
            school_name=school_name,
            school_folder_name=school_folder_name,
            worker=worker,
            mail_received_date=mail_received_date,
            work_date=work_date,
            open_date=open_date,
            roster_basis_date=roster_basis_date,
            scan=scan,
            result=result,
            note=note,
        ),
    )


def create_diff_tasklog(
    work_root: Path,
    *,
    school_name: str,
    school_folder_name: str,
    worker: str,
    mail_received_date: Optional[date],
    work_date: Optional[date],
    open_date: Optional[date],
    result: Any,
    note: str = "",
) -> TaskLogEntry:
    return append_new_tasklog_entry(
        work_root,
        entry_builder=lambda seq: build_diff_tasklog_entry(
            seq=seq,
            school_name=school_name,
            school_folder_name=school_folder_name,
            worker=worker,
            mail_received_date=mail_received_date,
            work_date=work_date,
            open_date=open_date,
            result=result,
            note=note,
        ),
    )


def list_school_progress_rows(
    work_root: Path,
    *,
    worker: str = "",
) -> List[Dict[str, str]]:
    rows = read_tasklog_rows(work_root)

    latest_by_school: Dict[str, Dict[str, str]] = {}

    def _seq_key(row: Dict[str, str]) -> int:
        try:
            return int(row.get("seq", 0))
        except Exception:
            return 0

    rows.sort(key=_seq_key)

    for row in rows:
        school_name = (row.get("school_name") or "").strip()
        if not school_name:
            continue

        if worker.strip():
            if (row.get("worker") or "").strip() != worker.strip():
                continue

        latest_by_school[school_name] = row

    out: List[Dict[str, str]] = []

    for school_name, row in sorted(latest_by_school.items(), key=lambda x: x[0]):
        status = _derive_school_progress_status(row)

        out.append({
            "school_name": school_name,
            "worker": row.get("worker", ""),
            "pipeline_kind": row.get("pipeline_kind", ""),
            "latest_seq": row.get("seq", ""),
            "latest_work_date": row.get("work_date", ""),
            "status": status,
            "bm_registered": row.get("bm_registered", ""),
            "email_sent": row.get("email_sent", ""),
            "sms_sent": row.get("sms_sent", ""),
            "extra_work_needed": row.get("extra_work_needed", ""),
            "extra_work_done": row.get("extra_work_done", ""),
            "note": row.get("note", ""),
            "error_summary": row.get("error_summary", ""),
        })

    return out


def _derive_school_progress_status(row: Dict[str, str]) -> str:
    error_summary = (row.get("error_summary") or "").strip()
    extra_work_needed = (row.get("extra_work_needed") or "").strip()
    extra_work_done = (row.get("extra_work_done") or "").strip()
    sms_sent = (row.get("sms_sent") or "").strip()
    email_sent = (row.get("email_sent") or "").strip()
    bm_registered = (row.get("bm_registered") or "").strip()
    file_created = (row.get("file_created") or "").strip()
    status = (row.get("status") or "").strip()

    if error_summary:
        return "오류"

    if extra_work_needed == "Y" and extra_work_done != "Y":
        return "추가 작업 필요"

    if sms_sent == "Y":
        return "문자 발송 완료"

    if email_sent == "Y":
        return "메일 발송 완료"

    if bm_registered == "Y":
        return "BM 등록 완료"

    if file_created == "Y":
        return "파일 생성 완료"

    if status:
        return status

    return "미작업"




