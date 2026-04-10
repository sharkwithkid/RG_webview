"""
app.py — ClassMate 진입점

실행:
    python app.py

의존성:
    pip install PyQt6 PyQt6-WebEngine
"""

import sys
import os
import ctypes
from pathlib import Path

from PyQt6.QtCore import QUrl, Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel

from bridge import Bridge

# ── 경로 기준 설정 ────────────────────────────────
# PyInstaller exe: 번들 리소스는 sys._MEIPASS, 사용자 데이터는 exe 옆
# 일반 실행: __file__ 기준
if getattr(sys, "frozen", False):
    _BUNDLE_DIR = Path(sys._MEIPASS)           # 번들 리소스 (읽기 전용)
    _APP_DIR    = Path(sys.executable).parent  # exe 옆 (설정·이력 저장용)
else:
    _BUNDLE_DIR = Path(__file__).parent
    _APP_DIR    = Path(__file__).parent

# core/ 모듈이 사용자 데이터 경로를 알 수 있도록 환경 변수로 공유
os.environ["RG_BUNDLE_DIR"] = str(_BUNDLE_DIR)
os.environ["RG_APP_DIR"]    = str(_APP_DIR)

# ── Windows High DPI 설정 ────────────────────────
# QApplication 생성 전에 설정해야 효과 있음
os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--enable-font-antialiasing --font-render-hinting=full")
QApplication.setHighDpiScaleFactorRoundingPolicy(
    Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
)


HTML_PATH = _BUNDLE_DIR / "ui" / "index.html"
ICON_PATH = _BUNDLE_DIR / "ClassMate.ico"
WINDOW_TITLE = "ClassMate"
APP_ID = "ReadingGate.ClassMate"

WINDOW_W, WINDOW_H = 930, 750
MIN_WINDOW_W, MIN_WINDOW_H = 930, 750
DEFAULT_ZOOM = 1.0


def _set_windows_appusermodelid(app_id: str) -> None:
    """작업표시줄 아이콘/그룹화에 쓰이는 Windows AppUserModelID 설정."""
    if os.name != "nt":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def _load_app_icon() -> QIcon:
    """번들/개발 환경 공통 아이콘 로드."""
    if ICON_PATH.exists():
        return QIcon(str(ICON_PATH))
    return QIcon()


def main():
    _set_windows_appusermodelid(APP_ID)

    app = QApplication(sys.argv)
    app.setApplicationName(WINDOW_TITLE)
    app.setDesktopFileName(APP_ID)
    # AA_UseHighDpiPixmaps는 PyQt6에서 제거됨 — 기본으로 활성화됨

    app_icon = _load_app_icon()
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    # ── Bridge 생성 ──────────────────────────────
    bridge = Bridge()

    # ── QWebChannel 등록 ─────────────────────────
    # JS에서 qt.webChannelTransport 로 접근 → "bridge" 객체 노출
    channel = QWebChannel()
    channel.registerObject("bridge", bridge)

    # ── WebEngineView 세팅 ───────────────────────
    view = QWebEngineView()
    view.page().setWebChannel(channel)
    view.setWindowTitle(WINDOW_TITLE)
    if not app_icon.isNull():
        view.setWindowIcon(app_icon)

    view.resize(WINDOW_W, WINDOW_H)
    view.setMinimumSize(MIN_WINDOW_W, MIN_WINDOW_H)
    view.setZoomFactor(DEFAULT_ZOOM)

    # ── index.html 로드 ──────────────────────────
    if not HTML_PATH.exists():
        print(f"[ERROR] index.html을 찾을 수 없습니다: {HTML_PATH}", file=sys.stderr)
        sys.exit(1)

    view.load(QUrl.fromLocalFile(str(HTML_PATH.resolve())))
    view.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
