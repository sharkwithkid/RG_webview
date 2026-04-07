"""
build.py — RG_webview 배포 빌드 스크립트

main 브랜치 push 시 git hook에서 자동 실행.
수동 실행: python build.py

순서:
  1. 브랜치 확인 (main 아니면 중단)
  2. 민감 파일 / 빌드 환경 점검
  3. PyInstaller 빌드
  4. Inno Setup 인스톨러 패키징
  5. 결과 출력
"""

import sys
import os
import subprocess
import json
import shutil
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).parent
DIST_DIR = ROOT / "dist" / "ClassMate"
OUTPUT_DIR = ROOT / "Output"
SPEC_FILE = ROOT / "RG_webview.spec"
ICON_FILE = ROOT / "ClassMate.ico"

# Inno Setup 기본 설치 경로 (없으면 스킵)
INNO_PATHS = [
    Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
    Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
]
ISS_FILE = ROOT / "ClassMate_installer.iss"

# ──────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────

def log(msg: str, level: str = "INFO"):
    prefix = {"INFO": "   ", "OK": " ✅", "WARN": " ⚠️", "ERR": " ❌", "HEAD": "\n🔧"}
    print(f"{prefix.get(level, '  ')} {msg}")

def abort(msg: str):
    log(msg, "ERR")
    log("빌드 중단.", "ERR")
    sys.exit(1)

def run(cmd: list, cwd=None, check=True):
    result = subprocess.run(cmd, cwd=cwd or ROOT, capture_output=True, text=True)
    if check and result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        abort(f"명령 실패: {' '.join(str(c) for c in cmd)}")
    return result

# ──────────────────────────────────────────────
# 1. 브랜치 확인
# ──────────────────────────────────────────────

def check_branch():
    log("브랜치 확인", "HEAD")
    result = run(["git", "rev-parse", "--abbrev-ref", "HEAD"])
    branch = result.stdout.strip()
    if branch != "release":
        abort(f"현재 브랜치: '{branch}'. 배포 빌드는 release 브랜치에서만 실행합니다.")
    log(f"브랜치: {branch}", "OK")

# ──────────────────────────────────────────────
# 2. 민감 파일 / 환경 점검
# ──────────────────────────────────────────────

def check_environment():
    log("환경 점검", "HEAD")

    # app_config.json — 실제 운영 데이터가 담겨있으면 배포본에 포함되면 안 됨
    config_path = ROOT / "app_config.json"
    if config_path.exists():
        try:
            cfg = json.loads(config_path.read_text(encoding="utf-8"))
            if cfg.get("worker_name") or cfg.get("work_root"):
                log("app_config.json에 운영 데이터가 있습니다. 배포본에서 제외됩니다.", "WARN")
        except Exception:
            pass

    # 불필요한 파일 체크
    dirty_files = ["run_error.log", "work_history_2026.json", "work_history_2025.json"]
    for f in dirty_files:
        if (ROOT / f).exists():
            log(f"{f} 발견 — 배포본에서 제외됩니다.", "WARN")

    # 필수 파일 체크
    for required in [SPEC_FILE, ICON_FILE]:
        if not required.exists():
            abort(f"필수 파일 없음: {required.name}")
    log("필수 파일 확인", "OK")

    # PyInstaller 설치 확인
    result = run(["pyinstaller", "--version"], check=False)
    if result.returncode != 0:
        abort("PyInstaller가 설치되어 있지 않습니다. pip install pyinstaller")
    log(f"PyInstaller {result.stdout.strip()}", "OK")

# ──────────────────────────────────────────────
# 3. PyInstaller 빌드
# ──────────────────────────────────────────────

def build_exe():
    log("PyInstaller 빌드", "HEAD")

    # 이전 빌드 정리
    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
        log("이전 dist 정리", "OK")

    build_dir = ROOT / "build"
    if build_dir.exists():
        shutil.rmtree(build_dir)

    result = subprocess.run(
        ["pyinstaller", str(SPEC_FILE), "--noconfirm"],
        cwd=ROOT,
        text=True,
    )
    if result.returncode != 0:
        abort("PyInstaller 빌드 실패. 위 로그를 확인하세요.")

    if not DIST_DIR.exists():
        abort(f"빌드 결과물을 찾을 수 없습니다: {DIST_DIR}")

    log(f"빌드 완료: {DIST_DIR}", "OK")

    # 배포본에서 민감 파일 제거
    for fname in ["app_config.json", "run_error.log"]:
        target = DIST_DIR / fname
        if target.exists():
            target.unlink()
            log(f"민감 파일 제거: {fname}", "OK")

    # app_config.example.json → app_config.json 으로 복사 (빈 설정으로 시작)
    example = ROOT / "app_config.example.json"
    if example.exists():
        shutil.copy(example, DIST_DIR / "app_config.json")
        log("app_config.example.json → app_config.json 복사", "OK")

# ──────────────────────────────────────────────
# 4. Inno Setup 인스톨러
# ──────────────────────────────────────────────

def build_installer():
    log("Inno Setup 인스톨러 빌드", "HEAD")

    iscc = next((p for p in INNO_PATHS if p.exists()), None)
    if not iscc:
        log("Inno Setup을 찾을 수 없습니다. 인스톨러 빌드를 건너뜁니다.", "WARN")
        log("설치: https://jrsoftware.org/isdl.php", "WARN")
        return

    if not ISS_FILE.exists():
        log(f"{ISS_FILE.name} 없음. 인스톨러 빌드를 건너뜁니다.", "WARN")
        return

    OUTPUT_DIR.mkdir(exist_ok=True)

    result = subprocess.run(
        [str(iscc), str(ISS_FILE)],
        cwd=ROOT,
        text=True,
    )
    if result.returncode != 0:
        abort("Inno Setup 빌드 실패.")

    installers = sorted(OUTPUT_DIR.glob("*.exe"))
    if installers:
        log(f"인스톨러 생성: {installers[-1].name}", "OK")
    else:
        log("인스톨러 파일을 찾을 수 없습니다.", "WARN")

# ──────────────────────────────────────────────
# 5. 결과 요약
# ──────────────────────────────────────────────

def summarize():
    log("빌드 완료", "HEAD")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    log(f"시각: {timestamp}", "INFO")

    exe = DIST_DIR / "ClassMate.exe"
    if exe.exists():
        size_mb = exe.stat().st_size / 1024 / 1024
        log(f"EXE: {exe}  ({size_mb:.1f} MB)", "OK")

    installers = sorted(OUTPUT_DIR.glob("*.exe")) if OUTPUT_DIR.exists() else []
    if installers:
        ins = installers[-1]
        size_mb = ins.stat().st_size / 1024 / 1024
        log(f"인스톨러: {ins}  ({size_mb:.1f} MB)", "OK")

# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

if __name__ == "__main__":
    check_branch()
    check_environment()
    build_exe()
    build_installer()
    summarize()
