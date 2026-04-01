"""
bridge.py — WebView ↔ Python 연결층

역할:
  JS(QWebChannel) 요청을 받아 engine/core를 호출하고 결과를 Signal로 돌려준다.
  비즈니스 로직 없음. 입력 검증 → 엔진 호출 → presenter 변환 → emit 이 전부.

슬롯 구성 (호출 순서 기준):
  A. 조회  — inspectWorkRoot, ensureWorkRootScaffold, loadSchoolNames,
              getSchoolDomain, getProjectDirs, loadNoticeTemplates,
              loadAppConfig, loadWorkHistory
  B. 저장  — saveAppConfig, saveWorkHistory, writeWorkResult, writeEmailSent
  C. 비동기 — startScanMain / startRunMain
              startScanDiff / startRunDiff
              startPreview
              (완료·실패는 pyqtSignal emit)
  D. OS연동 — pickWorkFolder, pickRosterLogFile, readXlsxMeta,
              openFile, openFolder, copyToClipboard

응답 포맷:
  동기  성공: { ok: true,  data: {...} }
  동기  실패: { ok: false, error: "메시지" }
  비동기 성공: { ok: true,  task: "scan_main", data: {...} }
  비동기 실패: { ok: false, task: "scan_main", error: "메시지", traceback: "..." }

날짜: 모든 날짜는 YYYY-MM-DD 문자열.
결과: ScanResult/PipelineResult 원본은 JS 미노출 — presenter 변환 후 전달.
"""

from __future__ import annotations

import time
import json
import re
import traceback as tb_module
from pathlib import Path

from PyQt6.QtCore import QObject, QThread, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import QApplication, QFileDialog

from engine import (
    inspect_work_root,
    scaffold_work_root,
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
from core.config_store import (
    load_app_config,
    save_app_config,
    load_work_history,
    save_work_history,
)
from core.presenter import (
    as_abs_issue_rows,
    present_diff_run_result,
    present_run_result,
    present_scan_result,
)



# ──────────────────────────────────────────────
# 공통 응답 헬퍼
# ──────────────────────────────────────────────

def ok_response(data: dict) -> str:
    return json.dumps({"ok": True, "data": data}, ensure_ascii=False)

def error_response(message: str, error_code: str = "") -> str:
    """동기 실패 응답. error_code는 선택사항 (이벤트 추적용)."""
    payload = {"ok": False, "error": message}
    if error_code:
        payload["error_code"] = error_code
    return json.dumps(payload, ensure_ascii=False)

def async_ok(task: str, data: dict) -> str:
    return json.dumps({"ok": True, "task": task, "data": data}, ensure_ascii=False)

def async_error(task: str, message: str, traceback: str = "", error_code: str = "") -> str:
    """비동기 실패 응답. error_code는 이벤트 추적용 (예: "SCAN_FAILED", "DATE_INVALID")."""
    payload = {"ok": False, "task": task, "error": message}
    if traceback:
        payload["traceback"] = traceback
    if error_code:
        payload["error_code"] = error_code
    return json.dumps(payload, ensure_ascii=False)



def _validate_date(date_str: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str or ""))


# ══════════════════════════════════════════════
# Workers  (비동기 작업 단위 — QThread 기반)
# ══════════════════════════════════════════════

class ScanWorker(QObject):
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def __init__(self, params: dict):
        super().__init__()
        self._params = params
        self.scan_result = None

    def run(self):
        try:

            started = time.time()

            result = scan_main_engine(
                work_root=self._params["work_root"],
                school_name=self._params["school_name"],
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
                school_kind_override=self._params.get("school_kind_override") or None,
            )


            self.scan_result = result
            payload = async_ok("scan_main", present_scan_result(result))
            self.finished.emit(payload)

        except Exception as e:
            self.scan_result = None
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
            self.finished.emit(async_ok("run_main", present_run_result(result)))
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
                target_year=self._params.get("target_year"),
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
            )
            self.finished.emit(async_ok("diff_scan", present_scan_result(result)))
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
                target_year=self._params.get("target_year"),
                school_start_date=self._params["school_start_date"],
                work_date=self._params["work_date"],
                roster_basis_date=self._params.get("roster_basis_date"),
                roster_xlsx=self._params.get("roster_xlsx") or None,
                col_map=self._params.get("col_map"),
                layout_overrides=self._params.get("layout_overrides"),
            )
            self.finished.emit(async_ok("diff_run", present_diff_run_result(result)))
        except Exception as e:
            self.failed.emit(async_error("diff_run", str(e), tb_module.format_exc()))


class PreviewWorker(QObject):
    """
    실제 시트의 1행부터 보여주는 미리보기 워커.
    시작행 이전은 JS에서 회색 처리하고, issue_rows는 경고 행으로 표시한다.
    """
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    MAX_PREVIEW_ROWS = 300

    def __init__(self, params: dict):
        super().__init__()
        self._params = params

    def run(self):
        kind = self._params.get("kind", "")
        try:
            from core.common import safe_load_workbook as _safe_wb
            from pathlib import Path as _Path

            file_path    = self._params["file_path"]
            sheet_name   = self._params.get("sheet_name", "")
            header_row   = int(self._params.get("header_row", 1))
            data_start   = int(self._params.get("data_start_row", 2))
            start_row    = int(self._params.get("start_row", 1) or 1)
            issue_rows   = list(self._params.get("issue_rows", []) or [])

            wb = _safe_wb(_Path(file_path), data_only=True, read_only=True)
            try:
                ws = (wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets[0])
                actual_sheet = getattr(ws, "title", None) or sheet_name

                hdr = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
                header_values = [str(c) if c is not None else "" for c in (hdr[0] if hdr else [])]

                max_cols = len(header_values)
                rows = []
                displayed_count = 0
                last_data_row = start_row - 1
                blank_streak = 0
                MAX_BLANK_STREAK = 10
                # 출력 파일(run_output)은 하나라도 값 있으면 데이터 행
                # 입력 파일은 핵심 열(이름/학년/반) 기준으로 판정
                is_output = (kind == "run_output")
                # 핵심 열 인덱스 추출 — 이름/학년/반 관련 헤더 위치
                KEY_KEYWORDS = {"이름", "성명", "학생이름", "학년", "반", "학급"}
                key_col_indices = [
                    i for i, h in enumerate(header_values)
                    if any(kw in str(h) for kw in KEY_KEYWORDS)
                ]
                header_col_count = max(len(header_values), 1)
                for row_idx, row in enumerate(ws.iter_rows(min_row=start_row, values_only=True), start=start_row):
                    vals = ["" if v is None else str(v) for v in row]
                    max_cols = max(max_cols, len(vals))
                    if is_output:
                        has_data = any(v.strip() for v in vals)
                    elif key_col_indices:
                        # 핵심 열 중 하나라도 값 있으면 데이터 행
                        has_data = any(vals[i].strip() for i in key_col_indices if i < len(vals))
                    else:
                        # 핵심 열 못 찾으면 과반수 기준 fallback
                        header_vals = vals[:header_col_count]
                        filled = sum(1 for v in header_vals if v.strip())
                        has_data = filled > (header_col_count / 2)
                    if has_data:
                        last_data_row = row_idx
                        blank_streak = 0
                    else:
                        blank_streak += 1
                        if blank_streak >= MAX_BLANK_STREAK:
                            break
                    if displayed_count < self.MAX_PREVIEW_ROWS:
                        rows.append(vals)
                        displayed_count += 1
                # 실제 데이터 행 기준으로 자르기 — 뒤쪽 빈/반빈 행 제거
                if last_data_row >= start_row:
                    rows = rows[:last_data_row - start_row + 1]
                    displayed_count = len(rows)
                # 실제 명단 수 = data_start_row 이후 행만 카운트 (예시행 제외)
                actual_count = max(0, last_data_row - data_start + 1)
                max_row = last_data_row

                columns = header_values if header_values else [""] * max_cols
                if len(columns) < max_cols:
                    columns += [""] * (max_cols - len(columns))
                rows = [r + [""] * (max_cols - len(r)) for r in rows]
                row_marks = {
                    'warn_rows': as_abs_issue_rows(issue_rows, data_start),
                    'error_rows': [],
                    'issue_rows': as_abs_issue_rows(issue_rows, data_start),
                    'muted_rows': list(range(start_row, max(data_start, start_row) )),
                }
            finally:
                wb.close()

            self.finished.emit(json.dumps({
                "ok": True, "kind": kind,
                "columns": columns, "rows": rows,
                "displayed_count": displayed_count,
                "actual_count": actual_count,
                "total_count": actual_count,
                "max_row": max_row,
                "truncated": bool(actual_count > displayed_count),
                "source_file": Path(file_path).name,
                "sheet_name": actual_sheet,
                "header_row": header_row,
                "data_start_row": data_start,
                "start_row": start_row,
                "issue_rows": as_abs_issue_rows(issue_rows, data_start),
                "row_marks": row_marks,
            }, ensure_ascii=False))

        except Exception as e:
            tb = tb_module.format_exc()
            self.failed.emit(json.dumps({
                "ok": False, "kind": kind,
                "error": str(e), "traceback": tb,
            }, ensure_ascii=False))


# ══════════════════════════════════════════════
# Bridge  (QWebChannel 등록 객체)
# 슬롯 순서: A.조회 → B.저장 → C.비동기 → D.OS연동
# ══════════════════════════════════════════════

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
    # A. 조회 계열 (동기, read-only, 부작용 없음)
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

    @pyqtSlot(str, result=str)
    def ensureWorkRootScaffold(self, work_root: str) -> str:
        """resources/templates/notices 폴더 생성 (없는 것만).
        반환: { scaffolded: ["resources", ...] }"""
        try:
            scaffolded = scaffold_work_root(work_root)
            return ok_response({"scaffolded": scaffolded})
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
            return ok_response({"config": load_app_config()})
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
            history = load_work_history(year)
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
            save_work_history(year, school_name, entry)
            return ok_response({})
        except Exception as e:
            return error_response(str(e))

    # ──────────────────────────────────────────
    # B. 저장 계열 (동기, write)
    # ──────────────────────────────────────────

    @pyqtSlot(str, result=str)
    def saveAppConfig(self, config_json: str) -> str:
        try:
            config = json.loads(config_json)
        except Exception:
            return error_response("잘못된 파라미터 형식입니다")
        try:
            save_app_config(config)
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
    # 시작: 즉시 ok_response 반환 → 완료/실패는 Signal emit
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
        if not _validate_date(params.get("school_start_date", "")):
            return error_response("개학일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
        if not _validate_date(params.get("work_date", "")):
            return error_response("작업일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
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
    # D. OS 연동 계열 (파일 선택·열기·클립보드)
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
            from core.common import safe_load_workbook as _safe_wb
            wb = _safe_wb(Path(str(xlsx_path)), data_only=True, read_only=True)

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
