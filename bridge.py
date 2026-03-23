"""
bridge.py — WebView ↔ Python 연결층

설계 원칙:
  - 호출/반환/상태 연결만 담당 (UI 로직·비즈니스 로직 금지)
  - ScanResult/RunResult 원본은 JS로 절대 넘기지 않음 → to_*_payload() 변환 후 전달
  - JS → Python: 슬롯 호출 + params_json
  - Python → JS: pyqtSignal emit (QWebChannel)
  - 모든 날짜는 YYYY-MM-DD 문자열로 통일

엔진 시그니처 주의사항 (실제 코드 기준):
  - load_all_school_names(roster_xlsx, col_map)  ← work_root 아님
  - get_school_domain(roster_xlsx, school_name, col_map)
  - get_project_dirs(work_root)  ← school_name 없음
  - run_main_engine(scan, work_date, school_start_date, ...)  ← ScanResult 객체 직접 받음
  - run_diff_engine(work_root, school_name, ...)  ← 내부에서 scan 재실행
"""

from __future__ import annotations

import json
import re
import traceback as tb_module
from pathlib import Path

from PyQt6.QtCore import QObject, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import QApplication, QFileDialog

from engine import (
    inspect_work_root,
    load_all_school_names,
    scan_main_engine,
    run_main_engine,
    scan_diff_engine,
    run_diff_engine,
    get_school_domain,
    get_project_dirs,
    load_notice_templates,
)
from core.roster_log import write_work_result, write_email_sent


# ── app_config 직접 읽기/쓰기 (별도 모듈 없음) ──────────────

CONFIG_PATH = Path("app_config.json")

DEFAULT_APP_CONFIG = {
    "work_root": "",
    "roster_log_path": "",
    "worker_name": "",
    "school_start_date": "",
    "work_date": "",
    "last_school": "",
    "roster_col_map": {
        "sheet": "",
        "header_row": 0,
        "data_start": 0,
        "col_school": 0,
        "col_email_arr": 0,
        "col_email_snt": 0,
        "col_worker": 0,
        "col_freshmen": 0,
        "col_transfer": 0,
        "col_withdraw": 0,
        "col_teacher": 0,
        "col_seq": 0,
    },
}

def _load_app_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return dict(DEFAULT_APP_CONFIG)

def _save_app_config(config: dict) -> None:
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ── work_history 직접 읽기/쓰기 ──────────────────────────────

def _work_history_path(school_year: int) -> Path:
    return Path(f"work_history_{school_year}.json")

def _load_work_history(school_year: int) -> dict:
    path = _work_history_path(school_year)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_work_history(school_year: int, school_name: str, entry: dict) -> None:
    history = _load_work_history(school_year)
    history[school_name] = entry
    _work_history_path(school_year).write_text(
        json.dumps(history, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ──────────────────────────────────────────────
# 공통 응답 헬퍼
# ──────────────────────────────────────────────

def ok_response(data: dict) -> str:
    return json.dumps({"ok": True, "data": data}, ensure_ascii=False)

def error_response(message: str) -> str:
    return json.dumps({"ok": False, "error": message}, ensure_ascii=False)

def async_ok(task: str, data: dict) -> str:
    return json.dumps({"ok": True, "task": task, "data": data}, ensure_ascii=False)

def async_error(task: str, message: str, traceback: str = "") -> str:
    return json.dumps(
        {"ok": False, "task": task, "error": message, "traceback": traceback},
        ensure_ascii=False,
    )


# ──────────────────────────────────────────────
# Result -> Payload 변환
# (ScanResult/PipelineResult 원본은 JS로 절대 노출하지 않는다)
# ──────────────────────────────────────────────

def _parse_log_entry(raw: str) -> dict:
    """'[INFO] ...' 형식 로그 문자열 -> {level, message} dict"""
    s = str(raw)
    if s.startswith("[ERROR]"):
        return {"level": "error", "message": s[7:].strip()}
    if s.startswith("[WARN]"):
        return {"level": "warn",  "message": s[6:].strip()}
    if s.startswith("[DEBUG]"):
        return {"level": "debug", "message": s[7:].strip()}
    if s.startswith("[INFO]"):
        return {"level": "info",  "message": s[6:].strip()}
    if s.startswith("[DONE]"):
        return {"level": "info",  "message": s[6:].strip()}
    return {"level": "info", "message": s.strip()}

def _logs_from_result(result) -> list:
    return [_parse_log_entry(l) for l in (getattr(result, "logs", None) or [])]


def to_scan_payload(result) -> dict:
    """ScanResult -> JS 전달용 요약 dict"""
    logs = _logs_from_result(result)
    warnings = [l for l in logs if l["level"] in ("warn", "error")]

    # ScanResult.freshmen / transfer_in / transfer_out / teachers 는
    # 각각 {file_name, sheet_name, header_row, data_start_row, ...} dict
    def _item(kind_key: str, kind_label: str):
        meta = getattr(result, kind_key, None)
        if meta is None:
            return None
        return {
            "kind":           kind_label,
            "file_name":      meta.get("file_name",      ""),
            "file_path":      meta.get("file_path",      ""),
            "sheet_name":     meta.get("sheet_name",     ""),
            "header_row":     meta.get("header_row",     1),
            "data_start_row": meta.get("data_start_row", 2),
            "row_count":      meta.get("row_count",      0),
        }

    items = [
        i for i in [
            _item("freshmen",    "신입생"),
            _item("transfer_in", "전입생"),
            _item("transfer_out","전출생"),
            _item("teachers",    "교직원"),
        ]
        if i is not None
    ]

    roster_basis_date = ""
    rbd = getattr(result, "roster_basis_date", None)
    if rbd is not None:
        roster_basis_date = rbd.isoformat() if hasattr(rbd, "isoformat") else str(rbd)

    # 학년도 아이디 규칙: roster_info.prefix_mode_by_roster_grade → {학년str: 연도int}
    grade_year_map = {}
    roster_info = getattr(result, "roster_info", None)
    if roster_info is not None:
        raw_map = getattr(roster_info, "prefix_mode_by_roster_grade", {}) or {}
        grade_year_map = {str(k): int(v) for k, v in raw_map.items()}

    return {
        "ok":                     bool(getattr(result, "ok", False)),
        "can_execute":            bool(getattr(result, "can_execute", False)),
        "can_execute_after_input":bool(getattr(result, "can_execute_after_input", False)),
        "missing_fields":         list(getattr(result, "missing_fields", []) or []),
        "needs_open_date":        bool(getattr(result, "needs_open_date", False)),
        "need_roster":            bool(getattr(result, "need_roster", False)),
        "roster_date_mismatch":   bool(getattr(result, "roster_date_mismatch", False)),
        "roster_basis_date":      roster_basis_date,
        "has_school_kind_warn":   False,
        "grade_year_map":         grade_year_map,
        "items":    items,
        "warnings": warnings,
        "logs":     logs,
    }


def to_run_payload(result) -> dict:
    """PipelineResult -> JS 전달용 요약 dict (Path -> str 전수 변환)"""
    logs = _logs_from_result(result)
    warnings = [l for l in logs if l["level"] in ("warn", "error")]

    audit    = getattr(result, "audit_summary", {}) or {}
    in_cnt   = audit.get("input_counts", {})

    return {
        "ok": bool(getattr(result, "ok", False)),
        "output_files": [
            {"name": p.name, "path": str(p)}
            for p in (getattr(result, "outputs", None) or [])
        ],
        "freshmen_count":          int(in_cnt.get("freshmen", 0)),
        "teacher_count":           int(in_cnt.get("teacher",  0)),
        "transfer_in_done":        int(getattr(result, "transfer_in_done",       0)),
        "transfer_in_hold":        int(getattr(result, "transfer_in_hold",       0)),
        "transfer_out_done":       int(getattr(result, "transfer_out_done",      0)),
        "transfer_out_hold":       int(getattr(result, "transfer_out_hold",      0)),
        "transfer_out_auto_skip":  int(getattr(result, "transfer_out_auto_skip", 0)),
        "warnings": warnings,
        "logs":     logs,
    }


def to_diff_run_payload(result) -> dict:
    """DiffPipelineResult -> JS 전달용 요약 dict"""
    logs = _logs_from_result(result)
    warnings = [l for l in logs if l["level"] in ("warn", "error")]

    return {
        "ok": bool(getattr(result, "ok", False)),
        "output_files": [
            {"name": p.name, "path": str(p)}
            for p in (getattr(result, "outputs", None) or [])
        ],
        "compare_only_count": int(getattr(result, "compare_only_count", 0)),
        "roster_only_count":  int(getattr(result, "roster_only_count",  0)),
        "matched_count":      int(getattr(result, "matched_count",      0)),
        "unresolved_count":   int(getattr(result, "unresolved_count",   0)),
        "transfer_in_done":   int(getattr(result, "transfer_in_done",   0)),
        "transfer_in_hold":   int(getattr(result, "transfer_in_hold",   0)),
        "transfer_out_done":  int(getattr(result, "transfer_out_done",  0)),
        "transfer_out_hold":  int(getattr(result, "transfer_out_hold",  0)),
        "warnings": warnings,
        "logs":     logs,
    }


def _validate_date(date_str: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str or ""))


# ──────────────────────────────────────────────
# Workers
# ──────────────────────────────────────────────

class ScanWorker(QObject):
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def __init__(self, params: dict):
        super().__init__()
        self._params = params
        self.scan_result = None  # Bridge가 run용으로 꺼내 쓴다

    def run(self):
        try:
            result = scan_main_engine(
                work_root=self._params["work_root"],
                school_name=self._params["school_name"],
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
            )
            self.scan_result = result  # 원본 보관 (Bridge만 접근)
            self.finished.emit(async_ok("scan_main", to_scan_payload(result)))
        except Exception as e:
            self.failed.emit(async_error("scan_main", str(e), tb_module.format_exc()))


class RunWorker(QObject):
    """run_main_engine은 ScanResult 객체를 직접 받는다."""
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def __init__(self, scan_result, work_date: str, school_start_date: str,
                 layout_overrides=None, school_kind_override=None):
        super().__init__()
        self._scan = scan_result
        self._work_date = work_date
        self._school_start_date = school_start_date
        self._layout_overrides = layout_overrides
        self._school_kind_override = school_kind_override

    def run(self):
        try:
            result = run_main_engine(
                scan=self._scan,
                work_date=self._work_date,
                school_start_date=self._school_start_date,
                layout_overrides=self._layout_overrides,
                school_kind_override=self._school_kind_override,
            )
            self.finished.emit(async_ok("run_main", to_run_payload(result)))
        except Exception as e:
            self.failed.emit(async_error("run_main", str(e), tb_module.format_exc()))


class DiffScanWorker(QObject):
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def __init__(self, params: dict):
        super().__init__()
        self._params = params

    def run(self):
        try:
            result = scan_diff_engine(
                work_root=self._params["work_root"],
                school_name=self._params["school_name"],
                target_year=int(self._params["target_year"]),
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
            )
            self.finished.emit(async_ok("diff_scan", to_scan_payload(result)))
        except Exception as e:
            self.failed.emit(async_error("diff_scan", str(e), tb_module.format_exc()))


class DiffRunWorker(QObject):
    """run_diff_engine은 내부에서 scan을 재실행하므로 scan 객체 불필요."""
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def __init__(self, params: dict):
        super().__init__()
        self._params = params

    def run(self):
        try:
            result = run_diff_engine(
                work_root=self._params["work_root"],
                school_name=self._params["school_name"],
                target_year=int(self._params["target_year"]),
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
            )
            self.finished.emit(async_ok("diff_run", to_diff_run_payload(result)))
        except Exception as e:
            self.failed.emit(async_error("diff_run", str(e), tb_module.format_exc()))


class PreviewWorker(QObject):
    """
    ScanResult 파일 메타(file_path, sheet_name, header_row, data_start_row)로
    실제 rows를 읽어 JS로 전달. 최대 MAX_PREVIEW_ROWS 행만 전달.
    """
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    MAX_PREVIEW_ROWS = 200

    def __init__(self, params: dict):
        super().__init__()
        self._params = params

    def run(self):
        kind = self._params.get("kind", "")
        try:
            from openpyxl import load_workbook as _load_wb

            file_path    = self._params["file_path"]
            sheet_name   = self._params.get("sheet_name", "")
            header_row   = int(self._params.get("header_row", 1))
            data_start   = int(self._params.get("data_start_row", 2))

            wb = _load_wb(file_path, data_only=True, read_only=True)
            try:
                ws = (wb[sheet_name]
                      if sheet_name and sheet_name in wb.sheetnames
                      else wb.worksheets[0])

                hdr = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
                columns = [str(c) if c is not None else "" for c in (hdr[0] if hdr else [])]

                all_rows = []
                for row in ws.iter_rows(min_row=data_start, values_only=True):
                    all_rows.append([str(c) if c is not None else "" for c in row])
                    if len(all_rows) >= self.MAX_PREVIEW_ROWS + 1:
                        break

                total_count = len(all_rows)
                truncated   = total_count > self.MAX_PREVIEW_ROWS
                rows        = all_rows[:self.MAX_PREVIEW_ROWS]
                actual_sheet = ws.title if hasattr(ws, "title") else sheet_name
            finally:
                wb.close()

            self.finished.emit(json.dumps({
                "ok": True, "kind": kind,
                "columns": columns, "rows": rows,
                "total_count": total_count, "truncated": truncated,
                "source_file":    Path(file_path).name,
                "sheet_name":     actual_sheet,
                "header_row":     header_row,
                "data_start_row": data_start,
            }, ensure_ascii=False))

        except Exception as e:
            self.failed.emit(json.dumps({
                "ok": False, "kind": kind,
                "error": str(e), "traceback": tb_module.format_exc(),
            }, ensure_ascii=False))


# ──────────────────────────────────────────────
# Bridge
# ──────────────────────────────────────────────

class Bridge(QObject):
    """
    WebView ↔ Python 연결 객체.
    QWebChannel에 등록하여 JS에서 직접 호출한다.
    """

    # Python -> JS 시그널
    scanFinished     = pyqtSignal(str)
    scanFailed       = pyqtSignal(str)
    runFinished      = pyqtSignal(str)
    runFailed        = pyqtSignal(str)
    diffScanFinished = pyqtSignal(str)
    diffScanFailed   = pyqtSignal(str)
    diffRunFinished  = pyqtSignal(str)
    diffRunFailed    = pyqtSignal(str)
    previewLoaded    = pyqtSignal(str)
    previewFailed    = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        # 중복 실행 방지 플래그
        self._is_scanning      = False
        self._is_running       = False
        self._is_diff_scanning = False
        self._is_diff_running  = False
        self._is_previewing    = False

        # 원본 결과 보관 (JS로 직접 노출 금지)
        # run_main_engine이 ScanResult를 직접 받으므로 Bridge에서 보관
        self._last_scan_result = None   # ScanResult (run_main_engine 인자용)
        self._last_run_result  = None   # PipelineResult
        self._current_output_files: list = []
        self._school_name_set: set = set()

        # Worker / Thread 참조 (GC 방지)
        self._scan_thread = None;       self._scan_worker = None
        self._run_thread  = None;       self._run_worker  = None
        self._diff_scan_thread = None;  self._diff_scan_worker = None
        self._diff_run_thread  = None;  self._diff_run_worker  = None
        self._preview_thread   = None;  self._preview_worker   = None

    def _start_worker(self, worker, thread, on_finished, on_failed):
        """Worker/Thread 연결 공통 처리"""
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(on_finished)
        worker.failed.connect(on_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    # ──────────────────────────────────────────
    # A. 조회 계열 (동기)
    # ──────────────────────────────────────────

    @pyqtSlot(str, result=str)
    def inspectWorkRoot(self, work_root: str) -> str:
        try:
            result = inspect_work_root(work_root)
            serializable = {
                k: str(v) if isinstance(v, Path) else v
                for k, v in result.items()
            }
            return ok_response(serializable)
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, str, result=str)
    def loadSchoolNames(self, roster_xlsx: str, col_map_json: str) -> str:
        """
        roster_xlsx: 명단 파일 경로 (work_root 아님)
        col_map_json: JSON 문자열 또는 "{}"
        """
        try:
            col_map = json.loads(col_map_json) if col_map_json else {}
        except Exception:
            col_map = {}
        try:
            names = load_all_school_names(
                roster_xlsx=Path(roster_xlsx) if roster_xlsx else None,
                col_map=col_map or None,
            )
            self._school_name_set = set(names)
            return ok_response({"school_names": names})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, str, str, result=str)
    def getSchoolDomain(self, roster_xlsx: str, school_name: str, col_map_json: str) -> str:
        try:
            col_map = json.loads(col_map_json) if col_map_json else {}
        except Exception:
            col_map = {}
        try:
            domain = get_school_domain(
                roster_xlsx=Path(roster_xlsx) if roster_xlsx else None,
                school_name=school_name,
                col_map=col_map or None,
            )
            return ok_response({"domain": domain or ""})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def getProjectDirs(self, work_root: str) -> str:
        """get_project_dirs(work_root) — school_name 파라미터 없음"""
        try:
            dirs = get_project_dirs(Path(work_root))
            return ok_response({"dirs": {k: str(v) for k, v in dirs.items()}})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def loadNoticeTemplates(self, work_root: str) -> str:
        try:
            templates = load_notice_templates(Path(work_root))
            return ok_response({"templates": templates})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(result=str)
    def loadAppConfig(self) -> str:
        try:
            return ok_response({"config": _load_app_config()})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def loadWorkHistory(self, school_year: str) -> str:
        """
        school_year: 연도 문자열 (예: "2025")
        반환: { history: { school_name: { last_date, worker, counts } } }
        """
        try:
            year = int(school_year)
        except (ValueError, TypeError):
            return error_response("잘못된 학년도 형식입니다")
        try:
            history = _load_work_history(year)
            return ok_response({"history": history})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, str, str, result=str)
    def saveWorkHistory(self, school_year: str, school_name: str, entry_json: str) -> str:
        """
        school_year: 연도 문자열
        school_name: 학교명
        entry_json: { last_date, worker, counts } JSON 문자열
        """
        try:
            year  = int(school_year)
            entry = json.loads(entry_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")
        try:
            _save_work_history(year, school_name, entry)
            return ok_response({})
        except Exception as e:
            return error_response(str(e))

    # ──────────────────────────────────────────
    # B. 저장 계열 (동기)
    # ──────────────────────────────────────────

    @pyqtSlot(str, result=str)
    def saveAppConfig(self, config_json: str) -> str:
        try:
            config = json.loads(config_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")
        try:
            _save_app_config(config)
            return ok_response({})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def writeWorkResult(self, params_json: str) -> str:
        """
        params: {
          xlsx_path, school_name, worker,
          kind_flags: {"신입생": bool, ...},
          email_arrived_date: "YYYY-MM-DD" | "",
          col_map: {...},
          seq_no: int | null
        }
        """
        try:
            p = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")
        try:
            from datetime import datetime as _dt
            arr_date = None
            raw = p.get("email_arrived_date", "")
            if raw and _validate_date(raw):
                arr_date = _dt.strptime(raw, "%Y-%m-%d").date()

            ok, msg = write_work_result(
                xlsx_path=Path(p["xlsx_path"]),
                school_name=p["school_name"],
                worker=p.get("worker", ""),
                kind_flags=p.get("kind_flags", {}),
                email_arrived_date=arr_date,
                col_map=p.get("col_map"),
                seq_no=p.get("seq_no"),
            )
            return ok_response({"message": msg}) if ok else error_response(msg)
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def writeEmailSent(self, params_json: str) -> str:
        """
        params: {
          xlsx_path, school_name,
          sent_date: "YYYY-MM-DD" | "",
          col_map: {...}
        }
        """
        try:
            p = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")
        try:
            from datetime import datetime as _dt
            sent_date = None
            raw = p.get("sent_date", "")
            if raw and _validate_date(raw):
                sent_date = _dt.strptime(raw, "%Y-%m-%d").date()

            ok, msg = write_email_sent(
                xlsx_path=Path(p["xlsx_path"]),
                school_name=p["school_name"],
                sent_date=sent_date,
                col_map=p.get("col_map"),
            )
            return ok_response({"message": msg}) if ok else error_response(msg)
        except Exception as e:
            return error_response(str(e))

    # ──────────────────────────────────────────
    # C. 비동기 시작 계열
    # ──────────────────────────────────────────

    @pyqtSlot(str, result=str)
    def startScanMain(self, params_json: str) -> str:
        """
        params: {
          work_root, school_name,
          school_start_date (YYYY-MM-DD), work_date (YYYY-MM-DD),
          roster_xlsx (optional), roster_basis_date (optional),
          col_map (optional)
        }
        완료 -> scanFinished(payload) / 실패 -> scanFailed(payload)
        """
        try:
            params = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")

        if not params.get("work_root"):
            return error_response("작업 폴더가 없습니다")
        if not params.get("school_name"):
            return error_response("학교가 선택되지 않았습니다")
        if not _validate_date(params.get("school_start_date", "")):
            return error_response("개학일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if not _validate_date(params.get("work_date", "")):
            return error_response("작업일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if self._is_scanning:
            return error_response("이미 스캔이 진행 중입니다")

        self._is_scanning = True
        self._last_scan_result = None

        worker = ScanWorker(params)
        thread = QThread()
        self._scan_worker = worker
        self._scan_thread = thread
        self._start_worker(worker, thread, self._on_scan_finished, self._on_scan_failed)
        return ok_response({})

    def _on_scan_finished(self, payload: str):
        self._is_scanning = False
        # Worker의 scan_result 원본을 Bridge에 보관 (run_main_engine 인자용)
        if self._scan_worker is not None:
            self._last_scan_result = getattr(self._scan_worker, "scan_result", None)
        self.scanFinished.emit(payload)

    def _on_scan_failed(self, payload: str):
        self._is_scanning = False
        self._last_scan_result = None
        self.scanFailed.emit(payload)

    # ── Run ────────────────────────────────────

    @pyqtSlot(str, result=str)
    def startRunMain(self, params_json: str) -> str:
        """
        run_main_engine은 ScanResult 객체를 직접 받는다.
        Bridge._last_scan_result (scanFinished 이후 자동 저장됨) 를 사용.

        params: {
          work_date (YYYY-MM-DD), school_start_date (YYYY-MM-DD),
          layout_overrides (optional), school_kind_override (optional)
        }
        완료 -> runFinished(payload) / 실패 -> runFailed(payload)
        """
        try:
            params = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")

        if self._last_scan_result is None:
            return error_response("스캔 결과가 없습니다. 먼저 스캔을 실행해 주세요.")
        if not getattr(self._last_scan_result, "ok", False):
            return error_response("스캔이 실패 상태입니다. 스캔을 다시 실행해 주세요.")
        if not _validate_date(params.get("work_date", "")):
            return error_response("작업일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if not _validate_date(params.get("school_start_date", "")):
            return error_response("개학일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if self._is_running:
            return error_response("이미 실행이 진행 중입니다")

        self._is_running = True

        worker = RunWorker(
            scan_result=self._last_scan_result,
            work_date=params["work_date"],
            school_start_date=params["school_start_date"],
            layout_overrides=params.get("layout_overrides"),
            school_kind_override=params.get("school_kind_override"),
        )
        thread = QThread()
        self._run_worker = worker
        self._run_thread = thread
        self._start_worker(worker, thread, self._on_run_finished, self._on_run_failed)
        return ok_response({})

    def _on_run_finished(self, payload: str):
        self._is_running = False
        self.runFinished.emit(payload)

    def _on_run_failed(self, payload: str):
        self._is_running = False
        self.runFailed.emit(payload)

    # ── Diff Scan ──────────────────────────────

    @pyqtSlot(str, result=str)
    def startScanDiff(self, params_json: str) -> str:
        """
        params: {
          work_root, school_name, target_year (int),
          school_start_date (YYYY-MM-DD), work_date (YYYY-MM-DD),
          roster_xlsx (optional), roster_basis_date (optional), col_map (optional)
        }
        """
        try:
            params = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")

        if not params.get("work_root"):
            return error_response("작업 폴더가 없습니다")
        if not params.get("school_name"):
            return error_response("학교가 선택되지 않았습니다")
        if not params.get("target_year"):
            return error_response("대상 연도가 없습니다")
        if not _validate_date(params.get("school_start_date", "")):
            return error_response("개학일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if not _validate_date(params.get("work_date", "")):
            return error_response("작업일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if self._is_diff_scanning:
            return error_response("이미 명단비교 스캔이 진행 중입니다")

        self._is_diff_scanning = True

        worker = DiffScanWorker(params)
        thread = QThread()
        self._diff_scan_worker = worker
        self._diff_scan_thread = thread
        self._start_worker(worker, thread, self._on_diff_scan_finished, self._on_diff_scan_failed)
        return ok_response({})

    def _on_diff_scan_finished(self, payload: str):
        self._is_diff_scanning = False
        self.diffScanFinished.emit(payload)

    def _on_diff_scan_failed(self, payload: str):
        self._is_diff_scanning = False
        self.diffScanFailed.emit(payload)

    # ── Diff Run ───────────────────────────────

    @pyqtSlot(str, result=str)
    def startRunDiff(self, params_json: str) -> str:
        """
        run_diff_engine은 내부에서 scan을 재실행 — scan 객체 불필요.
        params: {
          work_root, school_name, target_year (int),
          school_start_date (YYYY-MM-DD), work_date (YYYY-MM-DD),
          roster_basis_date (optional)
        }
        """
        try:
            params = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")

        if not params.get("work_root"):
            return error_response("작업 폴더가 없습니다")
        if not params.get("school_name"):
            return error_response("학교가 선택되지 않았습니다")
        if not params.get("target_year"):
            return error_response("대상 연도가 없습니다")
        if self._is_diff_running:
            return error_response("이미 명단비교 실행이 진행 중입니다")

        self._is_diff_running = True

        worker = DiffRunWorker(params)
        thread = QThread()
        self._diff_run_worker = worker
        self._diff_run_thread = thread
        self._start_worker(worker, thread, self._on_diff_run_finished, self._on_diff_run_failed)
        return ok_response({})

    def _on_diff_run_finished(self, payload: str):
        self._is_diff_running = False
        self.diffRunFinished.emit(payload)

    def _on_diff_run_failed(self, payload: str):
        self._is_diff_running = False
        self.diffRunFailed.emit(payload)

    # ── Preview ────────────────────────────────

    @pyqtSlot(str, result=str)
    def startPreview(self, params_json: str) -> str:
        """
        params: {
          kind: "freshmen"|"transfer_in"|"transfer_out"|"teachers"|"roster",
          file_path, sheet_name,
          header_row (int), data_start_row (int)
        }
        완료 -> previewLoaded(payload) / 실패 -> previewFailed(payload)
        """
        try:
            params = json.loads(params_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")

        if not params.get("kind"):
            return error_response("미리보기 종류가 지정되지 않았습니다")
        if not params.get("file_path"):
            return error_response("파일 경로가 없습니다")
        if self._is_previewing:
            return error_response("이미 미리보기가 로딩 중입니다")

        self._is_previewing = True

        worker = PreviewWorker(params)
        thread = QThread()
        self._preview_worker = worker
        self._preview_thread = thread
        self._start_worker(worker, thread, self._on_preview_loaded, self._on_preview_failed)
        return ok_response({})

    def _on_preview_loaded(self, payload: str):
        self._is_previewing = False
        self.previewLoaded.emit(payload)

    def _on_preview_failed(self, payload: str):
        self._is_previewing = False
        self.previewFailed.emit(payload)

    # ──────────────────────────────────────────
    # D. OS 연동 계열
    # ──────────────────────────────────────────

    @pyqtSlot(result=str)
    def pickWorkFolder(self) -> str:
        path = QFileDialog.getExistingDirectory(None, "작업 폴더 선택")
        return ok_response({"path": path or ""})

    @pyqtSlot(result=str)
    def pickRosterLogFile(self) -> str:
        path, _ = QFileDialog.getOpenFileName(
            None, "명단 파일 선택", "", "Excel 파일 (*.xlsx)"
        )
        return ok_response({"path": path or ""})

    @pyqtSlot(str, str, int, result=str)
    def readXlsxMeta(self, xlsx_path: str, sheet_name: str, header_row: int) -> str:
        """
        열 매핑 다이얼로그용 — xlsx 파일 시트 목록 + 지정 시트의 미리보기 반환.

        xlsx_path:  파일 경로
        sheet_name: 빈 문자열이면 첫 시트
        header_row: 헤더 행 번호 (1-based)

        반환:
          { sheets: [...], headers: [...], rows: [[...], ...] }
        """
        try:
            from openpyxl import load_workbook as _load_wb
            wb = _load_wb(str(xlsx_path), read_only=True, data_only=True)

            sheets = wb.sheetnames
            target = sheet_name if (sheet_name and sheet_name in sheets) else sheets[0]
            ws = wb[target]

            h_idx = max(0, header_row - 1)   # 0-based
            rows_raw = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows_raw.append([str(v) if v is not None else "" for v in row])
                if i >= h_idx + 15:           # 헤더 + 데이터 15행까지
                    break
            wb.close()

            headers = rows_raw[h_idx] if h_idx < len(rows_raw) else []
            preview  = rows_raw[h_idx + 1: h_idx + 11]  # 데이터 10행

            return ok_response({
                "sheets":  sheets,
                "sheet":   target,
                "headers": headers,
                "rows":    preview,
            })
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def openFile(self, path: str) -> str:
        import subprocess, sys, os
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
            return ok_response({})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def openFolder(self, path: str) -> str:
        import subprocess, sys, os
        try:
            target = str(Path(path).parent) if Path(path).is_file() else path
            if sys.platform == "win32":
                os.startfile(target)
            elif sys.platform == "darwin":
                subprocess.run(["open", target])
            else:
                subprocess.run(["xdg-open", target])
            return ok_response({})
        except Exception as e:
            return error_response(str(e))

    @pyqtSlot(str, result=str)
    def copyToClipboard(self, text: str) -> str:
        try:
            QApplication.clipboard().setText(text)
            return ok_response({})
        except Exception as e:
            return error_response(str(e))
