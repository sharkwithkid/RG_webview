"""
webview_app.py — WebView 버전 진입점

실행:
    python webview_app.py

의존성:
    pip install PyQt6 PyQt6-WebEngine

구버전(QWidget) 진입점은 app.py 그대로 유지.
"""

import sys
from pathlib import Path

from PyQt6.QtCore import QUrl
from PyQt6.QtWidgets import QApplication
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel

from bridge import Bridge


HTML_PATH = Path(__file__).parent / "ui" / "index.html"
WINDOW_TITLE = "리딩게이트 반이동 자동화"

WINDOW_W, WINDOW_H = 1040, 800
MIN_WINDOW_W, MIN_WINDOW_H = 1040, 800


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(WINDOW_TITLE)

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
    
    view.resize(WINDOW_W, WINDOW_H)
    view.setMinimumSize(MIN_WINDOW_W, MIN_WINDOW_H)
    #view.setMaximumSize(WINDOW_W, WINDOW_H)

    # ── index.html 로드 ──────────────────────────
    if not HTML_PATH.exists():
        print(f"[ERROR] index.html을 찾을 수 없습니다: {HTML_PATH}", file=sys.stderr)
        sys.exit(1)

    view.load(QUrl.fromLocalFile(str(HTML_PATH.resolve())))

    # dev_view = QWebEngineView()

    # view.page().setDevToolsPage(dev_view.page())
    # dev_view.setWindowTitle("DevTools")
    # dev_view.resize(1000, 700)
    # dev_view.show()
    
    view.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
