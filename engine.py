from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional, Union

from core.run_main import PipelineResult, run_pipeline, execute_pipeline
from core.run_diff import DiffPipelineResult, run_diff_pipeline
from core.scan_main import ScanResult, scan_pipeline, load_all_school_names, get_school_domain, scan_work_root, ensure_work_root_scaffold
from core.scan_diff import DiffScanResult, scan_diff_pipeline

from core.common import get_project_dirs, load_notice_templates

DateLike = Union[date, datetime, str]


def _to_path(value: Union[str, Path]) -> Path:
    return Path(value).resolve()


def _to_school_name(value: Optional[str]) -> str:
    return (value or "").strip()


def _to_date(value: DateLike) -> date:
    if isinstance(value, bool):
        raise TypeError(f"[ERROR] 날짜 값에 bool이 전달되었습니다. 날짜를 올바르게 설정해 주세요.")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        s = value.strip()
        if not s:
            raise ValueError("[ERROR] 날짜 값이 비어 있습니다.")
        return datetime.strptime(s, "%Y-%m-%d").date()
    raise TypeError(f"[ERROR] 지원하지 않는 날짜 타입입니다: {type(value).__name__}")


# =========================
# resources / work_root 점검
# =========================
def inspect_work_root(work_root: Union[str, Path]) -> Dict[str, Any]:
    """resources 폴더 구조 점검 — 순수 조회, 부작용 없음."""
    return scan_work_root(_to_path(work_root))


def scaffold_work_root(work_root: Union[str, Path]) -> list:
    """resources/templates/notices 폴더 생성. 이미 있으면 건너뜀.
    반환: 새로 생성된 폴더 이름 목록."""
    return ensure_work_root_scaffold(_to_path(work_root))


# =========================
# main pipeline
# =========================
def scan_main_engine(
    work_root: Union[str, Path],
    school_name: str,
    school_start_date: DateLike,
    work_date: DateLike,
    roster_basis_date: Optional[DateLike] = None,
    roster_xlsx: Optional[Union[str, Path]] = None,
    col_map: Optional[Dict[str, Any]] = None,
    school_kind_override: Optional[str] = None,
) -> ScanResult:
    """
    메인 반이동 작업의 사전 스캔만 수행.
    PyQt에서 '스캔/미리보기' 버튼에 연결하기 좋다.
    """
    work_root_p = _to_path(work_root)
    school_name_s = _to_school_name(school_name)
    roster_xlsx = Path(roster_xlsx).resolve() if roster_xlsx else None

    try:
        school_start_date_d = _to_date(school_start_date)
        work_date_d = _to_date(work_date)
        roster_basis_date_d = _to_date(roster_basis_date) if roster_basis_date is not None else None

        return scan_pipeline(
            work_root=work_root_p,
            school_name=school_name_s,
            school_start_date=school_start_date_d,
            work_date=work_date_d,
            roster_basis_date=roster_basis_date_d,
            roster_xlsx=roster_xlsx,
            col_map=col_map,
            school_kind_override=school_kind_override,
        )

    except Exception as e:
        return ScanResult(
            ok=False,
            logs=[f"[ERROR] {e}"],
            school_name=school_name_s,
            year_str=str(getattr(school_start_date, "year", "")) if not isinstance(school_start_date, str) else "",
            year_int=0,
            project_root=work_root_p,
        )


def run_main_engine(
    scan: ScanResult,
    work_date: DateLike,
    school_start_date: DateLike,
    layout_overrides: Optional[Dict[str, Dict[str, int]]] = None,
    school_kind_override: Optional[str] = None,
) -> PipelineResult:
    """
    메인 반이동 작업 실행용 엔진 진입점.
    scan_main_engine()의 결과(ScanResult)를 받아 execute_pipeline을 직접 호출.
    scan을 다시 돌리지 않으므로 중복 I/O 없음.
    """
    try:
        work_date_d = _to_date(work_date)
        school_start_date_d = _to_date(school_start_date)

        return execute_pipeline(
            scan=scan,
            work_date=work_date_d,
            school_start_date=school_start_date_d,
            layout_overrides=layout_overrides,
            school_kind_override=school_kind_override,
        )

    except Exception as e:
        return PipelineResult(
            ok=False,
            outputs=[],
            logs=[f"[ERROR] {e}"],
        )


# =========================
# diff pipeline
# =========================
def scan_diff_engine(
    work_root: Union[str, Path],
    school_name: str,
    target_year: Optional[int],
    school_start_date: DateLike,
    work_date: DateLike,
    roster_basis_date: Optional[DateLike] = None,
    roster_xlsx: Optional[Union[str, Path]] = None,
    col_map: Optional[Dict[str, Any]] = None,
) -> DiffScanResult:
    """
    diff 작업의 사전 스캔만 수행.
    PyQt에서 비교 전 파일 점검/레이아웃 확인용으로 사용.
    """
    work_root_p = _to_path(work_root)
    school_name_s = _to_school_name(school_name)
    roster_xlsx = Path(roster_xlsx).resolve() if roster_xlsx else None

    try:
        school_start_date_d = _to_date(school_start_date)
        work_date_d = _to_date(work_date)
        roster_basis_date_d = _to_date(roster_basis_date) if roster_basis_date is not None else None
        target_year_i = int(target_year) if target_year is not None else int(school_start_date_d.year)

        return scan_diff_pipeline(
            work_root=work_root_p,
            school_name=school_name_s,
            target_year=target_year_i,
            school_start_date=school_start_date_d,
            work_date=work_date_d,
            roster_basis_date=roster_basis_date_d,
            roster_xlsx=roster_xlsx,
            col_map=col_map,
        )

    except Exception as e:
        return DiffScanResult(
            ok=False,
            logs=[f"[ERROR] {e}"],
            school_name=school_name_s,
            year_str=str(target_year or ""),
            year_int=int(target_year) if target_year is not None and str(target_year).strip() else 0,
            project_root=work_root_p,
        )


def run_diff_engine(
    work_root: Union[str, Path],
    school_name: str,
    target_year: Optional[int],
    school_start_date: DateLike,
    work_date: DateLike,
    roster_basis_date: Optional[DateLike] = None,
    roster_xlsx: Optional[Union[str, Path]] = None,
    col_map: Optional[Dict[str, Any]] = None,
    layout_overrides: Optional[Dict[str, Dict[str, int]]] = None,
) -> DiffPipelineResult:
    """
    diff 작업 실행용 엔진 진입점.
    PyQt에서는 이 함수만 호출하면 된다.
    """
    work_root_p = _to_path(work_root)
    school_name_s = _to_school_name(school_name)
    roster_xlsx = Path(roster_xlsx).resolve() if roster_xlsx else None

    try:
        school_start_date_d = _to_date(school_start_date)
        work_date_d = _to_date(work_date)
        roster_basis_date_d = _to_date(roster_basis_date) if roster_basis_date is not None else None
        target_year_i = int(target_year) if target_year is not None else int(school_start_date_d.year)

        return run_diff_pipeline(
            work_root=work_root_p,
            school_name=school_name_s,
            target_year=target_year_i,
            school_start_date=school_start_date_d,
            work_date=work_date_d,
            roster_basis_date=roster_basis_date_d,
            roster_xlsx=roster_xlsx,
            col_map=col_map,
            layout_overrides=layout_overrides,
        )

    except Exception as e:
        return DiffPipelineResult(
            ok=False,
            outputs=[],
            logs=[f"[ERROR] {e}"],
        )


__all__ = [
    "inspect_work_root",
    "scaffold_work_root",
    "load_all_school_names",
    "scan_main_engine",
    "run_main_engine",
    "scan_diff_engine",
    "run_diff_engine",
    "get_school_domain",
    "get_project_dirs",
    "load_notice_templates",
]