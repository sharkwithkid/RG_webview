"""
run_dev.py — 개발용 핫리로드 실행기

ui/ 폴더의 파일 변경을 감지하면 앱을 자동으로 재시작합니다.

실행:
    python run_dev.py
"""

import sys
import time
import subprocess
import os
from pathlib import Path

WATCH_DIR = Path(__file__).parent / "ui"
WATCH_EXTS = {".html", ".js", ".css"}
APP_ENTRY = Path(__file__).parent / "app.py"
POLL_INTERVAL = 0.8  # 초


def get_mtimes():
    mtimes = {}
    for f in WATCH_DIR.rglob("*"):
        if f.suffix in WATCH_EXTS:
            try:
                mtimes[f] = f.stat().st_mtime
            except Exception:
                pass
    return mtimes


def main():
    print(f"[dev] 감시 폴더: {WATCH_DIR}")
    print(f"[dev] 변경 감지 대상: {', '.join(WATCH_EXTS)}")
    print(f"[dev] 앱 시작 중...\n")

    proc = subprocess.Popen([sys.executable, str(APP_ENTRY)])
    last_mtimes = get_mtimes()

    try:
        while True:
            time.sleep(POLL_INTERVAL)

            current_mtimes = get_mtimes()
            changed = [
                f for f, t in current_mtimes.items()
                if last_mtimes.get(f) != t
            ] + [
                f for f in last_mtimes if f not in current_mtimes
            ]

            if changed:
                for f in changed:
                    print(f"[dev] 변경 감지: {f.name}")
                print("[dev] 앱 재시작...\n")
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    proc.kill()

                proc = subprocess.Popen([sys.executable, str(APP_ENTRY)])
                last_mtimes = current_mtimes

            elif proc.poll() is not None:
                print("[dev] 앱이 종료되었습니다. 재시작 중...\n")
                proc = subprocess.Popen([sys.executable, str(APP_ENTRY)])
                last_mtimes = get_mtimes()

    except KeyboardInterrupt:
        print("\n[dev] 종료합니다.")
        proc.terminate()


if __name__ == "__main__":
    main()
