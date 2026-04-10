from __future__ import annotations

import json
import os
from copy import deepcopy
from pathlib import Path
from typing import Any

# ── 경로 기준 ────────────────────────────────────────────────────────
# PyInstaller exe 실행 시 app.py가 RG_APP_DIR 환경 변수를 설정함.
# 일반 실행 시에는 이 파일 기준 상위 폴더(앱 루트)를 사용.
def _app_dir() -> Path:
    env = os.environ.get("RG_APP_DIR")
    if env:
        return Path(env)
    return Path(__file__).resolve().parent.parent

BASE_DIR    = _app_dir()
CONFIG_PATH = BASE_DIR / "app_config.json"

DEFAULT_APP_CONFIG: dict[str, Any] = {
    'work_root': '',
    'roster_log_path': '',
    'worker_name': '',
    'school_start_date': '',
    'work_date': '',
    'last_school': '',
    'roster_col_map': {
        'sheet': '',
        'header_row': 0,
        'data_start': 0,
        'col_school': 0,
        'col_email_arr': 0,
        'col_email_snt': 0,
        'col_worker': 0,
        'col_freshmen': 0,
        'col_transfer': 0,
        'col_withdraw': 0,
        'col_teacher': 0,
        'col_seq': 0,
    },
}


def _deep_merge(base: Any, override: Any) -> Any:
    if isinstance(base, dict) and isinstance(override, dict):
        merged = {k: deepcopy(v) for k, v in base.items()}
        for key, value in override.items():
            if key in merged:
                merged[key] = _deep_merge(merged[key], value)
            else:
                merged[key] = deepcopy(value)
        return merged
    return deepcopy(override)


def load_app_config() -> dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            loaded = json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
            return _deep_merge(DEFAULT_APP_CONFIG, loaded)
        except Exception:
            pass
    return deepcopy(DEFAULT_APP_CONFIG)


def save_app_config(config: dict[str, Any]) -> None:
    normalized = _deep_merge(DEFAULT_APP_CONFIG, config or {})
    CONFIG_PATH.write_text(
        json.dumps(normalized, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )


def work_history_path(school_year: int) -> Path:
    return BASE_DIR / f'work_history_{school_year}.json'


def load_work_history(school_year: int) -> dict[str, Any]:
    path = work_history_path(school_year)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except Exception:
        return {}


def save_work_history(school_year: int, school_name: str, entry: dict[str, Any]) -> None:
    history = load_work_history(school_year)
    history[school_name] = entry
    work_history_path(school_year).write_text(
        json.dumps(history, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
