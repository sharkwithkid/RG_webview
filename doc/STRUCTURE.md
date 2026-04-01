# RG_webview 배포 구조도

## 최종 배포본

```
RG_webview_v20260401.zip  (3.15 MB, 36개 파일)
│
├── webview_app.py              # 앱 진입점 (PyQt6 + QWebEngineView)
├── bridge.py                   # JS ↔ Python 연결층 (QWebChannel)
│                               #   A. 조회 (read-only)
│                               #   B. 저장 (write)
│                               #   C. 비동기 시작 → Signal emit
│                               #   D. OS 연동 (파일·클립보드)
├── engine.py                   # 코어 공개 진입점
├── requirements.txt            # PyQt6, PyQt6-WebEngine, openpyxl
├── app_config.example.json     # 설정 파일 템플릿 (개인정보 없음)
├── RG_webview.spec             # PyInstaller exe 빌드 설정
│
├── core/
│   ├── events.py               # CoreEvent / RowMark 정의 (41개 케이스)
│   ├── presenter.py            # Result → UI payload 변환
│   ├── config_store.py         # app_config / work_history 저장소
│   ├── scan_main.py            # 반이동 스캔 파이프라인
│   ├── run_main.py             # 반이동 실행 파이프라인
│   ├── scan_diff.py            # 명단비교 스캔 파이프라인
│   ├── run_diff.py             # 명단비교 실행 파이프라인
│   ├── common.py               # 공통 유틸 (경로·헤더 감지·명부 로드)
│   ├── output_common.py        # 출력 파일 공통 유틸 (백업·셀 쓰기)
│   ├── roster_log.py           # 전체 명단 xlsx 읽기/쓰기
│   ├── xlsx_db.py              # 학교명/도메인 검색
│   └── utils.py                # 문자열 정규화
│
└── ui/
    ├── index.html              # 단일 HTML 진입점
    ├── main.js                 # 앱 bootstrap, 라우팅, 이벤트 바인딩
    ├── app_runtime.js          # Bridge Signal 처리, 워크플로우
    ├── app_state.js            # 전역 상태 (AppState 모듈)
    ├── ui_common.js            # 공통 렌더링 유틸 (UICommon)
    ├── scan_tab.js             # 스캔 탭
    ├── run_tab.js              # 실행 탭
    ├── diff_tab.js             # 명단비교 탭
    ├── status_panel.js         # 학교 선택 패널 + StatusUI
    ├── setup.js                # 초기 설정 화면
    ├── notice_tab.js           # 안내문 탭
    ├── col_map_dialog.js       # 열 매핑 다이얼로그
    ├── date_picker.js          # 날짜 선택기
    ├── qwebchannel.js          # Qt 제공 채널 클라이언트 (수정 금지)
    └── fonts/
        ├── Pretendard-Regular.woff2
        ├── Pretendard-Medium.woff2
        ├── Pretendard-SemiBold.woff2
        └── Pretendard-Bold.woff2
```

## 배포 방법

### 소스 배포 (Python 환경 필요)
```bash
pip install -r requirements.txt
python webview_app.py
```

### exe 배포 (팀원용, Windows에서 1회 빌드)
```bash
pip install pyinstaller
pip install -r requirements.txt
pyinstaller RG_webview.spec
# → dist/RG_webview/ 폴더를 팀원 PC에 복사
```

## 레이어 구조

```
[JS UI]  ←→  [bridge.py]  ←→  [engine.py]  ←→  [core/]
              QWebChannel        진입점 래퍼       비즈니스 로직
              Signal emit        타입 변환         순수 Python
```

- **core/**는 PyQt 의존 없음 — 단독 테스트 가능
- **bridge.py**는 입력 검증 + 엔진 호출 + presenter 변환만
- **UI**는 `status.messages` / `events` / `row_marks` 3가지만 보고 렌더링

## 실행 시 생성되는 파일 (배포본 미포함)

```
앱 루트/                        # exe의 경우 exe 옆 폴더
├── app_config.json             # 첫 실행 후 자동 생성, 설정 저장
└── work_history_YYYY.json      # 학교별 작업 이력
```
