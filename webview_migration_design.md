# WebView 전환 실행 설계서 (최종)

---

## 1. 상태 소유권 (Source of Truth)

### JS가 주 상태로 가져갈 것 (화면용 요약 데이터)

```
worker_name             작업자 이름
work_root               작업 폴더 경로
work_date               작업일 (YYYY-MM-DD)
school_start_date       개학일 (YYYY-MM-DD)
roster_log_path         명단 파일 경로
selected_school         현재 선택된 학교명
current_seq_no          자료실 순번
selected_domain         선택된 학교 도메인
school_names            전체 학교명 배열 (set 아님 — filter/includes로 충분)
pending_school_keyword  검색 대기 키워드
last_scan_logs          스캔 로그 (팝업용 요약)
last_run_logs           실행 로그 (팝업용 요약)
last_diff_logs          명단비교 로그 (팝업용 요약)
school_kind_active      학교 구분 활성 여부
school_kind_override    학교 구분 수동 설정값
pending_roster_log      명단 미기록 여부 (경고용)
pending_history_entry   저장 대기 중인 이력 항목
currentPage             "setup" | "main"
currentTab              "scan" | "run" | "notice"
isInitializing          앱 초기 로딩 중 여부
```

### Python이 주 상태로 가져갈 것 (실행용 원본 데이터)

```
_last_scan_result       ScanResult 원본 객체
_last_run_result        RunResult 원본 객체
_current_output_files   출력 파일 경로 리스트 (Path 객체)
_preview_rows           미리보기 대용량 rows 원본
_school_name_set        학교명 set (Python 내부 검색용 — JS 노출 안 함)
실행 중인 Worker 인스턴스들
```

### 동기화 규칙

- JS → Python: 브리지 함수 호출 시 params로 전달
- Python → JS: 이벤트 payload에 요약 데이터 포함해서 emit
- **ScanResult/RunResult 원본은 JS로 절대 넘기지 않는다** — Bridge에서 payload로 변환 후 전달
- JS에 있는 상태가 Python에 중복으로 존재해서는 안 된다
- 날짜는 항상 `YYYY-MM-DD` 문자열로 통일 (JS ↔ Python 포맷 불일치 방지)

---

## 2. Bridge API 분류

Bridge는 호출/반환/상태 연결층만 담당한다.

### Bridge에서 하지 말 것

```
❌ UI 조립 로직
❌ 복잡한 상태 파생 계산
❌ 화면별 표시 문구 생성
❌ 엔진 비즈니스 로직 재구현
❌ 엑셀 데이터 후처리 로직 중복 구현
```

이걸 어기면 JS에도 로직, Bridge에도 로직, 엔진에도 로직이 생겨 3중 구조 붕괴된다.

---

### A. 조회 계열 (즉시 반환, 동기)

```python
inspectWorkRoot(work_root)
loadSchoolNames(work_root)          # → 배열로 반환 (set 아님)
getSchoolDomain(work_root, school_name)
getProjectDirs(work_root, school_name)
loadNoticeTemplates(work_root)
loadAppConfig()
loadWorkHistory(school_year)
```

### B. 저장 계열 (즉시 반환, 동기)

```python
saveAppConfig(config_json)
saveWorkHistory(school_year, school_name, entry_json)
writeWorkResult(params_json)
writeEmailSent(params_json)
```

### C. 비동기 시작 계열 (시작만 하고 즉시 return)

```python
startScanMain(params_json)      # → scanFinished / scanFailed 시그널
startRunMain(params_json)       # → runFinished / runFailed 시그널
startScanDiff(params_json)      # → diffScanFinished / diffScanFailed 시그널
startRunDiff(params_json)       # → diffRunFinished / diffRunFailed 시그널
startPreview(params_json)       # → previewLoaded / previewFailed 시그널
```

### D. OS 연동 계열 (데스크톱 전용)

```python
pickWorkFolder()                # 폴더 선택 다이얼로그
pickRosterLogFile()             # 파일 선택 다이얼로그 (.xlsx)
openFile(path)                  # 파일 열기 (기본 앱으로)
openFolder(path)                # 탐색기로 폴더 열기
copyToClipboard(text)           # 클립보드 복사 (안내문 탭)
```

---

## 3. Bridge 입력 검증

Bridge는 엔진의 방어막 역할을 한다.
JS에서 잘못된 값이 들어오면 엔진 전에 잡아야 한다.

```python
@pyqtSlot(str, result=str)
def startScanMain(self, params_json: str) -> str:
    try:
        params = json.loads(params_json)
    except Exception:
        return error_response("잘못된 파라미터 형식입니다")

    if not params.get("work_root"):
        return error_response("작업 폴더가 없습니다")
    if not params.get("school_name"):
        return error_response("학교가 선택되지 않았습니다")
    if not params.get("school_start_date"):
        return error_response("개학일이 설정되지 않았습니다")
    if self._is_scanning:
        return error_response("이미 스캔이 진행 중입니다")

    self._is_scanning = True
    # Worker 시작...
    return ok_response({})
```

검증 대상 필수 항목:
- `work_root` 비어있음
- `school_name` 비어있음 / null
- 날짜 형식 이상 (YYYY-MM-DD 검사)
- 파일 경로 존재 여부
- 중복 실행 플래그 확인

---

## 4. 공통 응답 포맷

### 즉시 반환 (동기)

```json
{ "ok": true, "data": { ... } }
{ "ok": false, "error": "메시지" }
```

### 비동기 완료 이벤트 (Python → JS signal)

```json
{ "ok": true,  "task": "scan_main", "data": { ... } }
{ "ok": false, "task": "scan_main", "error": "메시지", "traceback": "..." }
```

비동기 실패 시 `traceback` 필드 포함 (디버깅용, 빈 문자열이어도 키는 항상 존재).

### 헬퍼 함수

```python
def ok_response(data: dict) -> str:
    return json.dumps({"ok": True, "data": data}, ensure_ascii=False)

def error_response(message: str) -> str:
    return json.dumps({"ok": False, "error": message}, ensure_ascii=False)

def async_ok(task: str, data: dict) -> str:
    return json.dumps({"ok": True, "task": task, "data": data}, ensure_ascii=False)

def async_error(task: str, message: str, tb: str = "") -> str:
    return json.dumps({"ok": False, "task": task, "error": message, "traceback": tb},
                      ensure_ascii=False)
```

---

## 5. 브리지 역호출 방식 (Python → JS)

`runJavaScript` 문자열 삽입 방식은 이스케이프 문제로 사용하지 않는다.
**QWebChannel + pyqtSignal emit 방식을 사용한다.**

```python
class Bridge(QObject):
    scanFinished      = pyqtSignal(str)   # JSON payload
    scanFailed        = pyqtSignal(str)
    runFinished       = pyqtSignal(str)
    runFailed         = pyqtSignal(str)
    diffScanFinished  = pyqtSignal(str)
    diffScanFailed    = pyqtSignal(str)
    diffRunFinished   = pyqtSignal(str)
    diffRunFailed     = pyqtSignal(str)
    previewLoaded     = pyqtSignal(str)
    previewFailed     = pyqtSignal(str)
```

```javascript
bridge.scanFinished.connect((payloadJson) => {
    const payload = JSON.parse(payloadJson)
    if (payload.ok) onScanFinished(payload.data)
    else onScanFailed(payload.error)
})
```

---

## 6. 비동기 이벤트 순서 보장

**규칙: 각 작업은 반드시 Started → (Finished | Failed) 순서로 발생한다.**

```
startScanMain 호출
    → 반드시 onScanStarted 먼저 발생
    → 이후 onScanFinished 또는 onScanFailed 중 정확히 하나만 발생
    → 둘 다 발생하거나 둘 다 안 오는 경우는 없다
```

이 규칙이 보장되어야 JS에서 race condition이 생기지 않는다.

### 전체 이벤트 목록

```
onScanStarted()                     스캔 시작
onScanFinished(data)                스캔 완료
onScanFailed(error, traceback)      스캔 실패

onRunStarted()
onRunFinished(data)
onRunFailed(error, traceback)

onDiffScanStarted()
onDiffScanFinished(data)
onDiffScanFailed(error, traceback)

onDiffRunStarted()
onDiffRunFinished(data)
onDiffRunFailed(error, traceback)

onPreviewStarted(kind)
onPreviewLoaded(kind, data)
onPreviewFailed(kind, error)
```

---

## 7. Worker 예외 처리

Worker 내부 예외는 반드시 잡아서 signal로 전달한다.
콘솔에만 찍히고 JS가 아무것도 못 받는 상황은 없어야 한다.

```python
class ScanWorker(QObject):
    finished = pyqtSignal(str)
    failed   = pyqtSignal(str)

    def run(self):
        try:
            result = scan_main_engine(...)
            self.finished.emit(async_ok("scan_main", to_scan_payload(result)))
        except Exception as e:
            import traceback
            self.failed.emit(async_error("scan_main", str(e), traceback.format_exc()))
```

---

## 8. 실행 중 플래그 (중복 실행 방지)

JS 상태에 아래 플래그를 유지한다.

```javascript
state = {
    isInitializing:    true,   // 앱 초기 로딩 중
    isScanning:        false,
    isRunning:         false,
    isDiffScanning:    false,
    isDiffRunning:     false,
    isPreviewLoading:  false,
}
```

### 규칙

```
시작 전  → 플래그 true 확인, true면 호출 차단
시작 시  → 플래그 = true, 버튼 비활성화
완료/실패 → 플래그 = false, 버튼 복원
```

취소 기능은 현재 단계에서 구현하지 않는다.
중복 실행 방지 + 로딩 상태 표시만으로 충분하다.

---

## 9. Payload 포맷 정의

### ScanResult → payload

```json
{
    "ok": true,
    "task": "scan_main",
    "data": {
        "summary": "신입생 45명, 전입생 3명, 전출생 2명, 교직원 12명",
        "warnings": [
            { "level": "warn", "message": "전출생 파일에 중복 항목이 있습니다" }
        ],
        "logs": [
            { "level": "info", "message": "신입생 명단 로드 완료" },
            { "level": "warn", "message": "..." }
        ],
        "items": [
            { "kind": "신입생", "file_name": "첨부2-1.xlsx", "row_count": 45 }
        ],
        "has_school_kind_warn": false
    }
}
```

### RunResult → payload

```json
{
    "ok": true,
    "task": "run_main",
    "data": {
        "summary": "반편성 완료 — 3개 파일 생성",
        "output_files": [
            { "name": "등록파일.xlsx",   "path": "C:\\...\\등록파일.xlsx" },
            { "name": "반이동파일.xlsx", "path": "C:\\...\\반이동파일.xlsx" }
        ],
        "logs": [
            { "level": "info", "message": "..." }
        ],
        "warnings": []
    }
}
```

### PreviewWorker → payload

```json
{
    "ok": true,
    "kind": "freshmen",
    "columns": ["학년", "반", "이름", "ID"],
    "rows": [ [...], [...] ],
    "total_count": 147,
    "truncated": true,
    "source_file": "첨부2-1.xlsx",
    "sheet_name": "신입생명단"
}
```

실패 시:
```json
{
    "ok": false,
    "kind": "freshmen",
    "error": "파일을 열 수 없습니다",
    "traceback": "..."
}
```

### 로그 레벨 구조

모든 로그/경고는 아래 구조를 따른다:
```json
{ "level": "info" | "warn" | "error", "message": "..." }
```

UI에서 색깔/아이콘 처리가 쉬워진다.

---

## 10. Result 객체 → Payload 변환

ScanResult, RunResult 원본은 JS로 넘기지 않는다.
Bridge에서 변환 함수를 통해 payload만 전달한다.
이 원칙은 절대 흔들리지 않는다.

```python
def to_scan_payload(result) -> dict:
    return {
        "summary": ...,
        "warnings": [...],
        "logs": [...],
        "items": [...],
        "has_school_kind_warn": bool(...),
    }

def to_run_payload(result) -> dict:
    return {
        "summary": ...,
        "output_files": [
            {"name": p.name, "path": str(p)}
            for p in result.outputs or []
        ],
        "logs": [...],
        "warnings": [...],
    }
```

**주의:** `Path` 객체는 전달 전 전부 `str()`로 변환한다.

---

## 11. 앱 초기화 시퀀스

```
1. JS 앱 시작 (DOMContentLoaded)
   → isInitializing = true, 로딩 화면 표시
        ↓
2. bridge.loadAppConfig()
        ↓
3. config.work_root 가 있으면 bridge.inspectWorkRoot(work_root)
   ├── 유효: 4번으로
   └── 무효: SetupPage 빈 상태 렌더링, isInitializing = false, 종료
        ↓
4. bridge.loadSchoolNames(work_root)
        ↓
5. config.last_school 있으면 복원 시도
   ├── 학교명 목록에 있음: selected_school 복원 → 6번으로
   └── 없음: 학교 선택 초기 상태
        ↓
6. 복원된 학교가 있으면:
   bridge.getSchoolDomain(work_root, school_name)
   bridge.getProjectDirs(work_root, school_name)
   bridge.loadNoticeTemplates(work_root)
        ↓
7. isInitializing = false, UI 렌더링
   (currentPage = "setup" or "main")
```

---

## 12. 화면/라우팅 구조

SPA 라우터는 초반에 넣지 않는다.
상태 기반 화면 전환으로 구현한다.

```javascript
state.currentPage = "setup" | "main"
state.currentTab  = "scan"  | "run" | "notice"
```

```
currentPage = "setup"
└── SetupPage
    작업 폴더, 명단 파일, 작업자, 날짜 입력

currentPage = "main"
├── StepBar (0~4 상태 표시)
├── StatusPanel (사이드바)
│   학교 검색, 현재 학교, 도착일/발송일 체크
└── MainTab
    currentTab = "scan"    스캔 실행 + 결과 테이블
    currentTab = "run"     실행 + 결과 파일 목록
    currentTab = "notice"  안내문 선택 + 복사
```

### StepBar 상태

```javascript
steps = [
    { label: "기본 설정", state: "done" | "active" | "idle" },
    { label: "학교 선택", state: "done" | "active" | "idle" },
    { label: "스캔",      state: "done" | "active" | "idle" | "warn" },
    { label: "실행·결과", state: "done" | "active" | "idle" },
    { label: "안내문",    state: "done" | "active" | "idle" },
]
```

---

## 13. 전환 순서 (권장)

```
1단계  bridge.py 골격 작성
       - 헬퍼 함수 (ok_response / error_response / async_ok / async_error)
       - to_scan_payload / to_run_payload 작성
       - ScanResult, RunResult → dict 변환 확인
       - Path → str 변환 전수 확인

2단계  SetupPage HTML 포팅
       - 가장 단순한 화면
       - pickWorkFolder / pickRosterLogFile OS 연동 확인

3단계  StatusPanel HTML 포팅
       - 학교 검색 자동완성
       - 상태 배지

4단계  MainTab 스캔 탭 포팅
       - 비동기 시작/완료/실패 이벤트 흐름 검증
       - 중복 실행 방지 플래그 검증
       - PreviewWorker 결과 테이블 렌더링

5단계  실행·결과 탭 포팅
       - output_files 목록 (name/path 구조)
       - openFile / openFolder OS 연동

6단계  안내문 탭 포팅
       - copyToClipboard

7단계  QWidget 코드 전체 제거
```

---

## 14. 건드리지 않을 것

```
engine.py 내부 로직
core/roster_log.py
Worker 클래스 (QThread 기반 그대로 유지)
app_config.json / work_history_*.json 포맷
ScanResult / RunResult 클래스 구조
```
