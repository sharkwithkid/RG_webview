# RG_webview — 개발자 레퍼런스

> 이 문서는 코드를 수정하거나 기능을 추가하는 개발자를 위한 레퍼런스입니다.  
> 코드가 암묵적으로 가정하고 있는 규칙, 레이어 설계, 확장 시 지켜야 할 원칙을 담습니다.

---

## 1. 배포 구조

```
RG_webview/
├── webview_app.py          # 앱 진입점 (PyQt6 + QWebEngineView)
├── bridge.py               # JS ↔ Python 연결층 (QWebChannel)
├── engine.py               # 코어 엔진 공개 진입점
├── requirements.txt        # 외부 의존성
├── app_config.example.json # 설정 파일 템플릿
├── build_release.py        # 배포 zip 생성 스크립트
│
├── core/                   # 비즈니스 로직 (PyQt 의존 없음)
│   ├── events.py           # CoreEvent / RowMark 정의 + 케이스별 생성 함수
│   ├── presenter.py        # Result → UI payload 변환
│   ├── config_store.py     # app_config / work_history 저장소
│   ├── scan_main.py        # 반이동 스캔 파이프라인
│   ├── run_main.py         # 반이동 실행 파이프라인
│   ├── scan_diff.py        # 명단비교 스캔 파이프라인
│   ├── run_diff.py         # 명단비교 실행 파이프라인
│   ├── common.py           # 공통 유틸 (경로, 헤더 감지, 명부 로드 등)
│   ├── output_common.py    # 출력 파일 공통 유틸 (백업, 셀 쓰기 등)
│   ├── roster_log.py       # 전체 명단 xlsx 읽기/쓰기
│   ├── xlsx_db.py          # 전체 명단에서 학교명/도메인 검색
│   ├── tasklog.py          # 작업 이력 CSV 기록
│   └── utils.py            # 문자열 정규화 유틸
│
└── ui/                     # 프론트엔드 (브라우저 환경)
    ├── index.html          # 단일 HTML 진입점
    ├── main.js             # 앱 bootstrap, 라우팅, 이벤트 바인딩
    ├── app_runtime.js      # Bridge signal 처리, 워크플로우 orchestration
    ├── app_state.js        # 전역 상태 객체 + AppState 모듈
    ├── ui_common.js        # 공통 렌더링 유틸 (UICommon)
    ├── scan_tab.js         # 스캔 탭
    ├── run_tab.js          # 실행 탭
    ├── diff_tab.js         # 명단비교 탭
    ├── status_panel.js     # 학교 선택 패널
    ├── setup.js            # 초기 설정 화면
    ├── notice_tab.js       # 안내문 탭
    ├── col_map_dialog.js   # 열 매핑 다이얼로그
    ├── date_picker.js      # 날짜 선택기
    ├── app_state.js        # 상태 관리
    ├── qwebchannel.js      # Qt WebChannel 클라이언트 (Qt 제공, 수정 금지)
    └── fonts/              # Pretendard 폰트
```

---

## 2. 레이어 설계 원칙

```
[코어] → [presenter] → [bridge] → [JS UI]
```

| 레이어 | 파일 | 해야 할 일 | 하면 안 되는 일 |
|---|---|---|---|
| **domain/core** | `scan_main`, `run_main`, `scan_diff`, `run_diff` | 판단·산출물 생성·이벤트 적재 | PyQt, HTML, 화면 카드 문구 조합 |
| **presenter** | `core/presenter.py` | Result → UI payload 변환, badge/status 계산 | 비즈니스 로직, I/O |
| **bridge** | `bridge.py` | 슬롯 수신 → 엔진 호출 → presenter → emit | 저장 로직(→ config_store), 변환 로직(→ presenter) |
| **UI** | `ui/*.js` | payload 받아서 렌더링 | `logs`로 UX 판정, `status.messages` 재해석 |

**핵심 원칙:**
- **코어는 사실을 만들고, presenter는 의미를 붙이고, UI는 보여주기만 한다.**
- UI는 반드시 `status.messages` / `events` / `row_marks` 3가지만 보고 판단한다. `logs`는 로그 다이얼로그 전용.
- 문구는 `events.py` 생성 함수에서만 만든다. bridge·JS에서 문구를 조합하지 않는다.

---

## 3. 폴더 구조 가정 (작업 폴더 기준)

코드가 암묵적으로 기대하는 구조:

```
작업폴더/  (work_root)
│
├── resources/              ← 이름에 'resources' 포함된 폴더 1개만 허용
│   ├── templates/          ← 등록/안내 템플릿 xlsx
│   │   ├── *등록*.xlsx     ← 파일명에 '등록' 포함, 정확히 1개
│   │   └── *안내*.xlsx     ← 파일명에 '안내' 포함, 정확히 1개
│   └── notices/            ← 안내문 텍스트 파일
│       └── *.txt           ← UTF-8 또는 UTF-8-SIG, 1개 이상
│
├── 학교명A/                 ← 학교 폴더 (이름 부분 일치로 탐색)
│   ├── *신입생*.xlsx        ← 파일명에 '신입생' 또는 '신입' 포함
│   ├── *전입생*.xlsx        ← 파일명에 '전입생' 또는 '전입' 포함
│   ├── *전출생*.xlsx        ← 파일명에 '전출생' 또는 '전출' 포함
│   ├── *교직원*.xlsx        ← 파일명에 '교사', '교원', '교직원' 포함
│   ├── *학생명부*.xlsx      ← 파일명에 '학생명부' 포함 (전입/전출 필요 시)
│   ├── *명렬표*.xlsx        ← 명단비교 탭용 재학생 파일 (키워드 우선 탐색)
│   └── 작업/               ← 출력 폴더 (자동 생성)
│       ├── ★{학교명}_등록작업파일(작업용).xlsx
│       ├── ☆{학교명}_{종류}_ID,PW안내.xlsx
│       ├── {학교명}_명단비교 결과.xlsx
│       └── _backup/        ← 덮어쓰기 시 이전 파일 자동 이동
│
└── 학교명B/
    └── ...
```

### 파일 탐색 세부 규칙

**resources 폴더**
- `work_root.iterdir()`에서 이름에 `"resources"` 포함된 폴더를 찾는다.
- 0개 → `work_root/resources`로 기본 경로 사용 (폴더 없으면 오류)
- 2개 이상 → `RESOURCES_CONFIG_ERROR` 발생

**학교 폴더**
- `work_root` 하위 폴더 중 이름에 학교명이 포함된 것을 찾는다 (`text_contains` 사용 — 공백·특수문자 무시 부분 일치).
- 0개 → `SCHOOL_FOLDER_NOT_FOUND`
- 2개 이상 → `SCHOOL_FOLDER_AMBIGUOUS`

**입력 파일 (.xlsx 한정)**
- 파일명에 키워드 포함 + `.xlsx` 확장자 + `~$`로 시작하지 않음
- `.xls` 감지 시 → `INPUT_XLS_FORMAT` 오류
- 같은 종류 2개 이상 → `DUPLICATE_INPUT_FILE` 오류
- 입력 파일은 학교 폴더 루트에 직접 위치해야 한다 (하위 폴더 탐색 없음)

**학생명부**
- 학교 폴더 루트에서 파일명에 `"학생명부"` 포함된 `.xlsx` 파일
- 여러 개면 최근 수정일 순으로 첫 번째 사용
- 필요한 경우: 전입 파일 있음 OR 전출 파일 있음 OR 신입생 파일에 1학년 외 학년 포함

**명단비교 재학생 파일**
- 1차: 키워드 (`명렬표`, `명렬`, `재학생`, `학생명단`) 포함 파일 수집
- 2차: 헤더 구조 검증 (`학년`, `반`, `이름` 열 존재 여부)
- 키워드 매칭 없으면 폴더 내 전체 xlsx에서 헤더 검증으로 fallback

---

## 4. 입력 파일 헤더 인식 규칙

헤더 행은 자동 감지한다 (위에서 아래로 스캔, 필수 슬롯이 모두 채워지는 행).

| 파일 종류 | 필수 열 | 인식 키워드 |
|---|---|---|
| 신입생 | name | 성명, 이름, 학생이름 |
| 전입생 | name | 성명, 이름 |
| 전출생 | name | 성명, 이름 |
| 교직원 | name | 성명, 이름, 성함, 교사명, 교원명, 교직원명, 선생님이름 등 |
| 명단비교 | name | 이름, 성명, 학생이름 |

공통 선택 열 키워드:

| 슬롯 | 키워드 |
|---|---|
| grade (학년) | 학년 |
| class (반) | 반, 학급 |
| learn (학습용 아이디 신청) | 학습용id신청, 학습용id, 학습용아이디 |
| admin (관리용 아이디 신청) | 관리용id신청, 관리용id, 관리용아이디 |

**데이터 시작 행**: 헤더 다음 줄부터 자동 감지. 예시 행(이름이 '예시', '샘플', '홍길동' 패턴)은 자동 제외.

---

## 5. 이벤트 시스템

### 원칙
- **모든 사용자 메시지는 `events.py` 생성 함수에서만 만든다.**
- 코어가 `CoreEvent`를 `result.events`에 적재 → `presenter.py`가 `status`로 변환 → UI가 렌더링.
- `logs`는 내부 디버그용. UI UX 판정에 사용 금지.

### CoreEvent 스키마

```python
@dataclass
class CoreEvent:
    code:       str           # 이벤트 식별자 (UPPER_SNAKE_CASE)
    level:      EventLevel    # "error" | "warn" | "hold" | "info"
    message:    str           # UI 카드에 표시할 한 줄 메시지
    detail:     str = ""      # 추가 설명 (필요 시)
    file_key:   FileKey = "global"  # "freshmen"|"transfer_in"|"transfer_out"|"teachers"|"roster"|"compare"|"global"
    row:        Optional[int] = None
    field_name: Optional[str] = None
    blocking:   bool = False  # True → 다음 단계 진행 불가
```

### level 의미

| level | 의미 | UI 표현 |
|---|---|---|
| `error` | 진행 불가 또는 결과 오류 | 빨간 배지/카드 |
| `warn` | 진행은 가능, 확인 필요 | 노란 배지/카드 |
| `hold` | 처리 완료되었으나 수동 확인 필요 | 분홍 배지/카드 |
| `info` | 참고 정보 | 카드 미표시 |

### 새 이벤트 추가 시

1. `events.py`에 생성 함수 추가
2. 코드명은 `UPPER_SNAKE_CASE`, 메시지는 사용자 언어
3. `file_key`를 정확히 지정
4. UI 카드에서 특별 처리가 필요하면 `blocking=True` 설정

---

## 6. Bridge API (JS → Python 슬롯)

모든 슬롯은 JSON 문자열을 주고받는다.

**응답 형식:**
```json
// 동기 슬롯 성공
{ "ok": true, "data": { ... } }

// 동기 슬롯 실패
{ "ok": false, "error": "메시지" }

// 비동기 슬롯 시작 응답 (실제 결과는 Signal로)
{ "ok": true, "data": {} }
```

**Signal 응답 형식 (비동기):**
```json
{ "ok": true, "task": "scan_main", "data": { ... } }
{ "ok": false, "task": "scan_main", "error": "메시지", "traceback": "..." }
```

### 동기 슬롯 목록

| 슬롯 | 설명 |
|---|---|
| `inspectWorkRoot(work_root)` | 작업 폴더 상태 점검 |
| `loadSchoolNames(roster_xlsx, col_map_json)` | 학교명 목록 로드 |
| `getSchoolDomain(roster_xlsx, school_name, col_map_json)` | 학교 도메인 조회 |
| `getProjectDirs(work_root)` | 프로젝트 폴더 경로 반환 |
| `loadNoticeTemplates(work_root)` | 안내문 템플릿 로드 |
| `loadAppConfig()` | 설정 파일 로드 |
| `saveAppConfig(config_json)` | 설정 파일 저장 |
| `loadWorkHistory(school_year)` | 작업 이력 로드 |
| `saveWorkHistory(school_year, school_name, entry_json)` | 작업 이력 저장 |
| `writeWorkResult(params_json)` | 전체 명단에 작업 결과 기록 |
| `writeEmailSent(params_json)` | 전체 명단에 이메일 발송 기록 |
| `pickWorkFolder()` | 폴더 선택 다이얼로그 |
| `pickRosterLogFile()` | 파일 선택 다이얼로그 |
| `readXlsxMeta(xlsx_path, sheet_name, header_row)` | 열 매핑 다이얼로그용 메타 |
| `openFile(path)` | 파일 열기 (OS 연동) |
| `openFolder(path)` | 폴더 열기 (OS 연동) |
| `copyToClipboard(text)` | 클립보드 복사 |

### 비동기 슬롯 + Signal 쌍

| 슬롯 | 완료 Signal | 실패 Signal |
|---|---|---|
| `startScanMain(params_json)` | `scanFinished` | `scanFailed` |
| `startRunMain(params_json)` | `runFinished` | `runFailed` |
| `startScanDiff(params_json)` | `diffScanFinished` | `diffScanFailed` |
| `startRunDiff(params_json)` | `diffRunFinished` | `diffRunFailed` |
| `startPreview(params_json)` | `previewLoaded` | `previewFailed` |

---

## 7. 설정 파일 (app_config.json)

`Path(__file__).resolve().parent.parent / 'app_config.json'` 위치에 저장.  
앱 루트 기준 고정 경로 → exe 배포 시에도 동일하게 동작.

```json
{
  "work_root": "",           // 작업 폴더 절대 경로
  "roster_log_path": "",     // 전체 명단 xlsx 절대 경로
  "worker_name": "",         // 작업자 이름
  "school_start_date": "",   // 개학일 (YYYY-MM-DD)
  "work_date": "",           // 작업일 (YYYY-MM-DD)
  "last_school": "",         // 마지막 선택 학교명
  "arrived_date": "",        // 이메일 도착일 (YYYY-MM-DD)
  "roster_col_map": {        // 전체 명단 열 매핑 (1-based)
    "sheet": "",
    "header_row": 0,
    "data_start": 0,
    "col_school": 0,
    "col_domain": 0,
    "col_email_arr": 0,
    "col_email_snt": 0,
    "col_worker": 0,
    "col_freshmen": 0,
    "col_transfer": 0,
    "col_withdraw": 0,
    "col_teacher": 0,
    "col_seq": 0
  }
}
```

**주의:** 배포본에 포함하지 않는다. `.gitignore`에 등록되어 있음.

---

## 8. 작업 이력 (work_history_YYYY.json)

앱 루트에 연도별로 생성. 배포본에 포함하지 않는다.

```json
{
  "학교명A": {
    "last_date": "2026-03-15",
    "worker": "홍길동",
    "counts": {
      "신입생": 42,
      "전입생": 3,
      "전출생": 1,
      "교직원": 8
    }
  }
}
```

---

## 9. 전체 명단 xlsx 열 매핑 규칙

`roster_log.py`와 `xlsx_db.py`가 공유하는 col_map 구조. 모든 열 번호는 1-based.

- `header_row`: 헤더 행 번호
- `data_start`: 데이터 시작 행 번호
- `col_school`: 학교명 열
- `col_domain`: 홈페이지/도메인 열 (없으면 `0` 또는 생략)
- `col_email_arr`: 이메일 도착일 열
- `col_email_snt`: 이메일 발송일 열
- `col_worker`: 담당자 열
- `col_freshmen` ~ `col_teacher`: 종류별 처리 결과 열
- `col_seq`: 자료실 순번 열 (없으면 `0` 또는 생략)

열 값 `0`은 `None` 취급 (미지정).

---

## 10. 출력 파일 규칙

| 파일 | 경로 | 명명 규칙 |
|---|---|---|
| 등록작업파일 | `학교폴더/작업/` | `★{학교명}_등록작업파일(작업용).xlsx` |
| 안내파일 | `학교폴더/작업/` | `☆{학교명}_{종류}_ID,PW안내.xlsx` |
| 명단비교 결과 | `학교폴더/작업/` | `{학교명}_명단비교 결과.xlsx` |
| 백업 | `학교폴더/작업/_backup/` | `{원본명}_{YYYYmmdd_HHMMSS}.xlsx` |

- **출력 폴더(`작업/`)는 자동 생성**된다.
- 동일 파일명이 이미 있으면 `_backup/`으로 이동 후 덮어쓴다.

---

## 11. 코드 수정 시 체크리스트

**새 이벤트 케이스 추가:**
- [ ] `events.py`에 생성 함수 추가
- [ ] 코어 파이프라인에서 발행 (`sr.events.append(...)`)
- [ ] `doc/error_case_v4.xlsx` 매트릭스 업데이트

**새 입력 파일 종류 추가:**
- [ ] `FRESHMEN_KEYWORDS` 류에 키워드 추가 (`scan_main.py`)
- [ ] `HEADER_SLOTS` 딕셔너리에 슬롯 추가 (`common.py`)
- [ ] `ScanResult`에 해당 필드 추가
- [ ] `presenter.py`의 `present_scan_result`에 `normalize_scan_item` 호출 추가
- [ ] `bridge.py` `PreviewWorker`에 kind 처리 추가

**UI 공통 렌더링 변경:**
- [ ] `ui_common.js`만 수정 → scan/run/diff 탭 자동 반영
- [ ] 탭별 개별 렌더링 로직은 해당 `*_tab.js`만 수정

---

## 12. 개발 환경 실행

```bash
# 의존성 설치
pip install -r requirements.txt

# 개발 모드 (UI 파일 변경 시 자동 재시작)
python run_dev.py

# 배포 zip 생성
python build_release.py
```
