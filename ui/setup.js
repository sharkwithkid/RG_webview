/**
 * setup.js — SetupPage 로직
 *
 * 의존: main.js (state, bridge, App, _el, _todayStr)
 * HTML ID: work-root, roster-log, worker, school-start-date, work-date,
 *           btn-start, col-map-badge, work-root-badge, setup-banner,
 *           stat-school, stat-date, stat-step
 */

'use strict';

const Setup = (() => {

  let _isStarting = false;

  // ──────────────────────────────────────────────
  // 초기화 (main.js initApp에서 호출)
  // ──────────────────────────────────────────────
  function init(cfg) {
    cfg = cfg || {};

    _setVal('work-root',          cfg.work_root          || '');
    _setVal('roster-log',         cfg.roster_log_path    || '');
    _setVal('worker',             cfg.worker_name        || '');
    _setVal('school-start-date',  cfg.school_start_date  || _todayStr());
    _setVal('work-date',          _todayStr());   // 작업일은 항상 오늘

    // roster_col_map 상태 반영
    const cm = cfg.roster_col_map || {};
    state.roster_col_map = cm;
    _updateColMapBadge(cm);

    // work-root 배지 초기화
    _updateWorkRootBadge(cfg.work_root ? 'ok' : null, cfg.work_root ? '폴더 확인 완료 ✓' : '');

    refreshBtn();

    // 이벤트 연결 (중복 방지: 한 번만)
    if (!_el('work-root')._setupBound) {
      _el('work-root').addEventListener('change', onWorkRootChange);
      _el('work-root').addEventListener('input',  refreshBtn);
      _el('worker').addEventListener('input',     refreshBtn);
      _el('school-start-date').addEventListener('input', e => _autoHyphenDate(e.target));
      _el('work-date').addEventListener('input',         e => _autoHyphenDate(e.target));
      _el('work-root')._setupBound = true;
    }
  }

  // ──────────────────────────────────────────────
  // 버튼 활성화 조건: work_root + worker
  // ──────────────────────────────────────────────
  function refreshBtn() {
    const ok = !!_getVal('work-root') && !!_getVal('worker');
    _el('btn-start').disabled = !ok;
  }

  // ──────────────────────────────────────────────
  // 작업 폴더 선택
  // ──────────────────────────────────────────────
  async function pickFolder() {
    const res = JSON.parse(await bridge.pickWorkFolder());
    if (!res.ok || !res.data.path) return;
    _setVal('work-root', res.data.path);
    refreshBtn();
    await onWorkRootChange();
  }

  // ──────────────────────────────────────────────
  // 작업 폴더 변경 시 즉시 점검
  // ──────────────────────────────────────────────
  async function onWorkRootChange() {
    const path = _getVal('work-root');
    if (!path) { _updateWorkRootBadge(null); return; }

    // 1) 폴더 구조 생성 (write) — scaffold 전용 API, inspect와 분리
    const scaffoldRes = JSON.parse(await bridge.ensureWorkRootScaffold(path));
    if (scaffoldRes.ok) {
      const scaffolded = scaffoldRes.data.scaffolded || [];
      if (scaffolded.length) {
        toast(
          `작업 폴더 구조를 생성했습니다.\n${scaffolded.map(f => '  • ' + f).join('\n')}\n\ntemplates 폴더에 등록·안내 템플릿 파일을,\nnotices 폴더에 안내문 txt 파일을 넣어 주세요.`,
          'info', 6000
        );
      }
    }

    // 2) 순수 상태 조회 (read-only) — 부작용 없음
    const res = JSON.parse(await bridge.inspectWorkRoot(path));
    if (!res.ok) {
      _updateWorkRootBadge('err', '폴더 오류');
      return;
    }
    if (!res.data.ok) {
      const msgs = (res.data.errors || [])
        .map(e => e.replace(/^\[ERROR\]\s*/, ''))
        .join('\n');
      _updateWorkRootBadge('err', '구성 오류', msgs);
      _el('btn-start').disabled = true;
    } else {
      _updateWorkRootBadge('ok', '폴더 확인 완료 ✓');
    }
  }

  // ──────────────────────────────────────────────
  // 명단 파일 선택
  // ──────────────────────────────────────────────
  async function pickRoster() {
    const res = JSON.parse(await bridge.pickRosterLogFile());
    if (!res.ok || !res.data.path) return;

    const xlsxPath = res.data.path;

    // xls 파일 선택 시 즉시 차단
    if (xlsxPath.toLowerCase().endsWith('.xls') && !xlsxPath.toLowerCase().endsWith('.xlsx')) {
      _showBanner('err', '명단 파일은 .xlsx 형식이어야 합니다. .xls 파일은 Excel에서 .xlsx로 저장한 뒤 다시 선택해 주세요.');
      return;
    }

    _setVal('roster-log', xlsxPath);

    // 열 매핑 다이얼로그 (콜백 방식)
    await ColMap.open(xlsxPath, state.roster_col_map || {}, _onColMapResult);
  }

  // 열 매핑 완료 콜백
  async function _onColMapResult(resultMap) {
    state.roster_col_map = resultMap;
    _updateColMapBadge(resultMap);

    // 즉시 저장 (기존 config 보존)
    const cfgRes   = JSON.parse(await bridge.loadAppConfig());
    const existing = cfgRes.ok ? (cfgRes.data.config || {}) : {};
    const cfg = Object.assign({}, existing, {
      roster_log_path: _getVal('roster-log'),
      roster_col_map:  resultMap,
    });
    const saveRes = JSON.parse(await bridge.saveAppConfig(JSON.stringify(cfg)));
    if (saveRes.ok) {
      _showBanner('ok', '명단 파일 열 설정이 저장되었습니다.');
      _autoClear();
    } else {
      _showBanner('err', '저장 실패: ' + saveRes.error);
    }
  }

  // ──────────────────────────────────────────────
  // 기본 설정 저장
  // ──────────────────────────────────────────────
  async function saveDefaults() {
    // 기존 config를 먼저 읽어 roster_col_map 등 보존
    const cfgRes = JSON.parse(await bridge.loadAppConfig());
    const existing = cfgRes.ok ? (cfgRes.data.config || {}) : {};

    const cfg = Object.assign({}, existing, {
      work_root:         _getVal('work-root'),
      roster_log_path:   _getVal('roster-log'),
      worker_name:       _getVal('worker'),
      school_start_date: _getVal('school-start-date'),
      work_date:         _getVal('work-date'),
    });

    const res = JSON.parse(await bridge.saveAppConfig(JSON.stringify(cfg)));
    if (res.ok) {
      _showBanner('ok', '기본 설정이 저장되었습니다.');
      _autoClear();
    } else {
      _showBanner('err', '저장 실패: ' + res.error);
    }
  }

  // ──────────────────────────────────────────────
  // 기본 설정 불러오기
  // ──────────────────────────────────────────────
  async function loadDefaults() {
    const res = JSON.parse(await bridge.loadAppConfig());
    if (!res.ok) { _showBanner('err', '불러오기 실패: ' + res.error); return; }
    init(res.data.config || {});
    _showBanner('ok', '기본 설정을 불러왔습니다.');
    _autoClear();
  }

  // ──────────────────────────────────────────────
  // 작업 시작
  // ──────────────────────────────────────────────
  async function start() {
    if (_isStarting) return;

    const workRoot  = _getVal('work-root');
    const worker    = _getVal('worker');
    const rosterLog = _getVal('roster-log');

    if (!workRoot)  { _showBanner('err', '작업 폴더를 입력하세요.');          return; }
    if (!worker)    { _showBanner('err', '작업자 이름을 입력하세요.');         return; }
    if (!rosterLog) { _showBanner('err', '학교 전체 명단 파일을 선택하세요.'); return; }

    const cm = state.roster_col_map || {};
    if (!cm.col_school) {
      _showBanner('warn', '명단 파일의 열 매핑을 완료해야 합니다.\n찾아보기로 파일을 선택하고 열 매핑을 설정하세요.');
      return;
    }
    if (!cm.col_domain) {
      _showBanner('warn', '명단 파일에서 도메인(홈페이지) 열이 지정되지 않았습니다.');
      return;
    }

    _isStarting = true;
    const btn = _el('btn-start');
    btn.innerHTML = '<span class="spinner"></span>확인 중...';
    btn.disabled  = true;

    try {
      // 1) resources 점검
      const inspRes = JSON.parse(await bridge.inspectWorkRoot(workRoot));
      if (!inspRes.ok) { _showBanner('err', '폴더 오류: ' + inspRes.error); return; }
      if (!inspRes.data.ok) {
        const msgs = (inspRes.data.errors || [])
          .map(e => e.replace(/^\[ERROR\]\s*/, ''))
          .join('\n');
        _showBanner('warn', 'resources 구성 확인 필요:\n' + msgs);
        return;
      }

      // 2) state 업데이트
      state.work_root         = workRoot;
      state.roster_log_path   = rosterLog;
      state.worker_name       = worker;
      state.school_start_date = _getVal('school-start-date');
      state.work_date         = _getVal('work-date');
      state.roster_col_map    = cm;

      // 3) main.js onSetupComplete 호출
      await App.onSetupComplete({
        work_root:         state.work_root,
        roster_log_path:   state.roster_log_path,
        worker_name:       state.worker_name,
        school_start_date: state.school_start_date,
        work_date:         state.work_date,
      });

    } finally {
      _isStarting = false;
      btn.textContent = '작업 시작 →';
      refreshBtn();
    }
  }

  // ──────────────────────────────────────────────
  // 열 매핑 badge 업데이트
  // ──────────────────────────────────────────────
  function _updateColMapBadge(cm) {
    const badge = _el('col-map-badge');
    if (!badge) return;
    const mapped = !!(cm && cm.col_school && cm.col_domain);
    badge.textContent = mapped ? '열 매핑 완료 ✓' : '열 매핑 필요';
    badge.className   = `file-badge ${mapped ? 'ok' : 'warn'}`;
  }

  // ──────────────────────────────────────────────
  // 작업 폴더 badge 업데이트
  // ──────────────────────────────────────────────
  function _updateWorkRootBadge(type, label, tooltip) {
    const badge = _el('work-root-badge');
    if (!badge) return;
    if (!type) {
      badge.textContent = '';
      badge.className   = 'file-badge';
      badge.title       = '';
      return;
    }
    badge.textContent = label;
    badge.className   = `file-badge ${type}`;
    badge.title       = tooltip || '';
  }

  // ──────────────────────────────────────────────
  // 배너 헬퍼 → toast() 위임
  // ──────────────────────────────────────────────
  function _showBanner(type, msg) {
    // type 매핑: 'ok'→'ok', 'warn'→'warn', 'err'→'err'
    toast(msg, type === 'err' ? 'err' : type === 'warn' ? 'warn' : 'ok');
  }

  function _hideBanner() { /* toast는 자동 소멸 */ }

  function _autoClear() { /* toast는 자동 소멸 */ }

  // ──────────────────────────────────────────────
  // 날짜 자동 하이픈 (YYYYMMDD → YYYY-MM-DD)
  // ──────────────────────────────────────────────
  function _autoHyphenDate(el) {
    let v = el.value.replace(/[^0-9]/g, '');
    if (v.length > 4) v = v.slice(0, 4) + '-' + v.slice(4);
    if (v.length > 7) v = v.slice(0, 7) + '-' + v.slice(7);
    el.value = v.slice(0, 10);
  }

  // ──────────────────────────────────────────────
  // DOM 헬퍼
  // ──────────────────────────────────────────────
  function _getVal(id) { return (_el(id)?.value || '').trim(); }
  function _setVal(id, v) { const el = _el(id); if (el) el.value = v; }

  // ──────────────────────────────────────────────
  // Public
  // ──────────────────────────────────────────────
  return { init, refreshBtn, pickFolder, onWorkRootChange, pickRoster, saveDefaults, loadDefaults, start };

})();
