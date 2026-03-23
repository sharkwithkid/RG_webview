/**
 * main.js — 앱 상태, 라우팅, Bridge 연결, 이벤트 바인딩
 *
 * 설계 원칙:
 *  - JS 주 상태: state 객체
 *  - inline onclick 없음 — _bindEvents()에서 모든 이벤트 등록
 *  - alert 없음 — toast() 사용
 *  - innerHTML 최소화 — 동적 생성 요소는 createElement 사용
 */

'use strict';

// ──────────────────────────────────────────────
// 전역 앱 상태 (JS Source of Truth)
// ──────────────────────────────────────────────
const state = {
  worker_name:        '',
  work_root:          '',
  work_date:          '',
  school_start_date:  '',
  roster_log_path:    '',
  roster_col_map:     {},
  arrived_date:       '',
  school_folders:     [],   // work_root 내 학교 폴더명 목록 (예: ["270. 용인정평초", ...])

  school_names:       [],
  selected_school:    '',
  selected_domain:    '',
  current_seq_no:     null,

  last_scan_logs:     [],
  last_run_logs:      [],
  last_diff_logs:     [],

  isInitializing:     true,
  isScanning:         false,
  isRunning:          false,
  isDiffScanning:     false,
  isDiffRunning:      false,
  isPreviewLoading:   false,

  currentPage:        'setup',
  currentTab:         'scan',

  pending_roster_log: false,
};

// ──────────────────────────────────────────────
// Bridge
// ──────────────────────────────────────────────
let bridge = null;

// ──────────────────────────────────────────────
// Toast 시스템 (alert 대체)
// ──────────────────────────────────────────────
function toast(msg, type = 'info', duration = 3000) {
  // type: 'ok' | 'warn' | 'err' | 'info'
  const container = _getOrCreateToastContainer();
  const el = document.createElement('div');
  el.style.cssText = `
    padding: 10px 16px; border-radius: 8px; font-size: 13px; font-weight: 600;
    max-width: 380px; word-break: keep-all; line-height: 1.5;
    box-shadow: 0 4px 16px rgba(0,0,0,.12);
    animation: toast-in .2s ease;
  `;
  const styles = {
    ok:   'background:#DCFCE7;border:1px solid #BBF7D0;color:#15803D',
    warn: 'background:#FEF9C3;border:1px solid #FDE047;color:#92400E',
    err:  'background:#FEE2E2;border:1px solid #FECACA;color:#DC2626',
    info: 'background:#DBEAFE;border:1px solid #BFDBFE;color:#1D4ED8',
  };
  el.style.cssText += ';' + (styles[type] || styles.info);
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    el.style.transition = 'opacity .3s';
    setTimeout(() => el.remove(), 300);
  }, duration);
}

function _getOrCreateToastContainer() {
  let c = document.getElementById('toast-container');
  if (!c) {
    c = document.createElement('div');
    c.id = 'toast-container';
    c.style.cssText = 'position:fixed;bottom:24px;right:24px;z-index:9999;display:flex;flex-direction:column;gap:8px;align-items:flex-end';
    // toast-in 애니메이션
    const style = document.createElement('style');
    style.textContent = '@keyframes toast-in{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}';
    document.head.appendChild(style);
    document.body.appendChild(c);
  }
  return c;
}

// ──────────────────────────────────────────────
// 초기화 시퀀스
// ──────────────────────────────────────────────
async function initApp() {
  try {
    bridge = await _connectBridge();
    _connectSignals();

    const cfgRes = JSON.parse(await bridge.loadAppConfig());
    if (!cfgRes.ok) throw new Error('설정 로드 실패: ' + cfgRes.error);
    const cfg = cfgRes.data.config || {};

    Setup.init(cfg);

    if (cfg.work_root) {
      const inspRes = JSON.parse(await bridge.inspectWorkRoot(cfg.work_root));
      if (!inspRes.ok || !inspRes.data.ok) {
        state.isInitializing = false;
        _showPage('setup');
        return;
      }

      state.work_root         = cfg.work_root;
      state.worker_name       = cfg.worker_name       || '';
      state.school_start_date = cfg.school_start_date || '';
      state.work_date         = cfg.work_date         || _todayStr();
      state.roster_log_path   = cfg.roster_log_path   || '';
      state.roster_col_map    = cfg.roster_col_map    || {};
      state.arrived_date      = cfg.arrived_date      || '';

      if (cfg.roster_log_path && cfg.roster_col_map?.col_school) {
        const namesRes = JSON.parse(
          await bridge.loadSchoolNames(
            cfg.roster_log_path,
            JSON.stringify(cfg.roster_col_map || {})
          )
        );
        if (namesRes.ok) {
          state.school_names = namesRes.data.school_names || [];
          Panel.init(state.school_names);
        }
      }
    }

  } catch (e) {
    console.error('initApp error:', e);
  } finally {
    state.isInitializing = false;
    _showPage('setup');
    _bindEvents();
    DatePicker.init();
    DatePicker._syncAll();
  }
}

// ──────────────────────────────────────────────
// 이벤트 바인딩 (inline onclick 없음)
// ──────────────────────────────────────────────
function _bindEvents() {

  // ── data-action 버튼 위임 ─────────────────────
  document.addEventListener('click', e => {
    const btn = e.target.closest('[data-action]');
    if (btn) {
      e.stopPropagation();
      _handleAction(btn.dataset.action, btn);
    }

    // 스핀 버튼 (data-spin-kind / data-spin-delta)
    const spin = e.target.closest('[data-spin-kind]');
    if (spin) {
      Scan.spin(spin.dataset.spinKind, parseInt(spin.dataset.spinDelta, 10));
    }

    // 필터 버튼 (data-filter)
    const filter = e.target.closest('[data-filter]');
    if (filter) {
      Scan.toggleFilter(filter.dataset.filter);
    }

    // 스텝바 클릭 (data-step)
    const step = e.target.closest('.step-item[data-step]');
    if (step) {
      App.goStep(parseInt(step.dataset.step, 10));
    }
  });

  // ── input 이벤트 ──────────────────────────────
  _el('work-root').addEventListener('input',  () => Setup.refreshBtn());
  _el('work-root').addEventListener('change', () => Setup.onWorkRootChange());
  _el('worker').addEventListener('input',     () => Setup.refreshBtn());

  _el('school-input').addEventListener('input',   e => Panel.onInput(e.target.value));
  _el('school-input').addEventListener('keydown', e => Panel.onKeyDown(e));

  _el('preview-search').addEventListener('input', () => Scan.filterPreview());
  _el('run-search').addEventListener('input',     () => Run.filterTable());

  // arrived-date 변경 시 state에 저장 (다음 실행 시 복원)
  _el('arrived-date').addEventListener('change', e => {
    state.arrived_date = e.target.value || '';
    bridge.saveAppConfig(JSON.stringify(_readFullConfig()));
  });

  // col_map dialog
  _el('cm-sheet').addEventListener('change',     () => ColMap.onSheetChange());
  _el('cm-header-row').addEventListener('change',() => ColMap.onHeaderRowChange());

  // ── 이벤트 위임: 동적 요소 ───────────────────
  // output-file-list (클릭으로 파일 열기) → run_tab.js에서 직접 처리
  // school-dropdown → status_panel.js에서 직접 처리
  // notice-list → notice_tab.js에서 직접 처리
}

function _handleAction(action, el) {
  const actions = {
    // Setup
    'pick-folder':    () => Setup.pickFolder(),
    'pick-roster':    () => Setup.pickRoster(),
    'save-defaults':  () => Setup.saveDefaults(),
    'load-defaults':  () => Setup.loadDefaults(),
    'start':          () => Setup.start(),

    // Mode / nav
    'mode-main':      () => App.setMode('main'),
    'mode-diff':      () => App.setMode('diff'),
    'goto-run':       () => Scan.goToRun(),
    'goto-notice':    () => App.goStep(4),

    // Panel
    'apply-school':   () => Panel.apply(),
    'toggle-grade':   () => Panel.toggleGrade(),
    'apply-grade':    () => Panel.applyGrade(),
    'record-roster':  () => Panel.recordRoster(),
    'open-roster':    () => Panel.openRoster(),
    'new-school':     () => Panel.newSchool(),

    // Scan
    'scan-start':     () => Scan.start(),
    'scan-log':       () => Scan.showLog(),
    'toggle-viewer':  () => Scan.toggleViewer(),

    // Run
    'run-start':      () => Run.start(),
    'run-log':        () => Run.showLog(),
    'run-dup':        () => Run.toggleDup(),
    'open-folder':    () => Run.openFolder(),

    // Notice
    'notice-copy':    () => Notice.copy(),
    'notice-reset':   () => Notice.reset(),

    // Diff
    'diff-start':     () => Diff.start(),
    'diff-log':       () => Diff.showLog(),

    // ColMap dialog
    'cm-skip':        () => ColMap.skipRole(),
    'cm-cancel':      () => ColMap.cancel(),
    'cm-confirm':     () => ColMap.confirm(),
  };
  if (actions[action]) actions[action]();
}

// ──────────────────────────────────────────────
// QWebChannel 연결
// ──────────────────────────────────────────────
function _connectBridge() {
  return new Promise(resolve => {
    if (typeof QWebChannel === 'undefined') {
      resolve(_makeMockBridge());
      return;
    }
    new QWebChannel(qt.webChannelTransport, ch => resolve(ch.objects.bridge));
  });
}

// ──────────────────────────────────────────────
// Bridge 시그널 연결
// ──────────────────────────────────────────────
function _connectSignals() {
  bridge.scanFinished.connect(payload => {
    const p = JSON.parse(payload);
    state.isScanning     = false;
    state.last_scan_logs = p.data?.logs || [];
    if (p.ok) Scan.onFinished(p.data);
    else      Scan.onFailed(p.error || '스캔 실패');
  });
  bridge.scanFailed.connect(payload => {
    const p = JSON.parse(payload);
    state.isScanning     = false;
    state.last_scan_logs = [{ level: 'error', message: p.error || '' }];
    Scan.onFailed(p.error || '예기치 못한 오류');
  });
  bridge.runFinished.connect(payload => {
    const p = JSON.parse(payload);
    state.isRunning     = false;
    state.last_run_logs = p.data?.logs || [];
    if (p.ok) Run.onFinished(p.data);
    else      Run.onFailed(p.error || '실행 실패');
  });
  bridge.runFailed.connect(payload => {
    const p = JSON.parse(payload);
    state.isRunning     = false;
    state.last_run_logs = [{ level: 'error', message: p.error || '' }];
    Run.onFailed(p.error || '예기치 못한 오류');
  });
  bridge.diffScanFinished.connect(payload => {
    const p = JSON.parse(payload);
    state.isDiffScanning = false;
    if (p.ok) Diff.onScanFinished(p.data);
    else      Diff.onScanFailed(p.error || '명단 비교 스캔 실패');
  });
  bridge.diffScanFailed.connect(payload => {
    const p = JSON.parse(payload);
    state.isDiffScanning = false;
    Diff.onScanFailed(p.error || '예기치 못한 오류');
  });
  bridge.diffRunFinished.connect(payload => {
    const p = JSON.parse(payload);
    state.isDiffRunning  = false;
    state.last_diff_logs = p.data?.logs || [];
    if (p.ok) Diff.onFinished(p.data);
    else      Diff.onFailed(p.error || '명단 비교 실패');
  });
  bridge.diffRunFailed.connect(payload => {
    const p = JSON.parse(payload);
    state.isDiffRunning  = false;
    state.last_diff_logs = [{ level: 'error', message: p.error || '' }];
    Diff.onFailed(p.error || '예기치 못한 오류');
  });
  bridge.previewLoaded.connect(payload => {
    const p = JSON.parse(payload);
    state.isPreviewLoading = false;
    if (p.ok) {
      if (p.kind === 'run_output') Run.onPreviewLoaded(p);
      else                         Scan.onPreviewLoaded(p);
    } else {
      if (p.kind === 'run_output') _el('run-preview-info').textContent = p.error || '미리보기 실패';
      else                         Scan.onPreviewFailed(p.kind, p.error || '미리보기 실패');
    }
  });
  bridge.previewFailed.connect(payload => {
    const p = JSON.parse(payload);
    state.isPreviewLoading = false;
    Scan.onPreviewFailed(p.kind || '', p.error || '예기치 못한 오류');
  });
}

// ──────────────────────────────────────────────
// App 네임스페이스
// ──────────────────────────────────────────────
const App = {

  async onSetupComplete(params) {
    state.work_root         = params.work_root;
    state.roster_log_path   = params.roster_log_path;
    state.worker_name       = params.worker_name;
    state.school_start_date = params.school_start_date;
    state.work_date         = params.work_date;

    await bridge.saveAppConfig(JSON.stringify(_readFullConfig()));

    _el('header-work-date').textContent = `작업일 · ${state.work_date}`;

    // 학교 폴더 목록 로드
    const inspRes = JSON.parse(await bridge.inspectWorkRoot(state.work_root));
    state.school_folders = inspRes.ok ? (inspRes.data.school_folders || []) : [];

    const namesRes = JSON.parse(
      await bridge.loadSchoolNames(
        state.roster_log_path,
        JSON.stringify(state.roster_col_map || {})
      )
    );
    state.school_names = namesRes.ok ? (namesRes.data.school_names || []) : [];
    Panel.init(state.school_names);
    Panel.setWorkContext({ work_date: state.work_date, arrived_date: state.arrived_date });

    App.setStepState(0, 'done');
    for (let i = 1; i <= 4; i++) App.setStepState(i, 'idle');

    _showPage('main');
    App.goTab('scan');
    App.setMode('main');
  },

  async onSchoolSelected(schoolName) {
    state.selected_school = schoolName;
    state.current_seq_no  = null;

    const domRes = JSON.parse(
      await bridge.getSchoolDomain(
        state.roster_log_path,
        schoolName,
        JSON.stringify(state.roster_col_map || {})
      )
    );
    state.selected_domain = domRes.ok ? (domRes.data.domain || '') : '';

    const tplRes = JSON.parse(await bridge.loadNoticeTemplates(state.work_root));
    if (tplRes.ok) Notice.loadTemplates(tplRes.data.templates || {}, _noticeCtx());

    // 작업 이력 로드
    const schoolYear = (state.work_date || _todayStr()).slice(0, 4);
    const histRes = JSON.parse(await bridge.loadWorkHistory(schoolYear));
    const histEntry = histRes.ok ? (histRes.data.history?.[schoolName] || null) : null;

    let history_text = null;
    if (histEntry) {
      const countStr = Object.entries(histEntry.counts || {})
        .filter(([, v]) => v)
        .map(([k, v]) => `${k} ${v}명`)
        .join(' · ');
      history_text = `마지막 작업 · ${histEntry.last_date || '-'}`;
      if (histEntry.worker) history_text += ` (${histEntry.worker})`;
      if (countStr) history_text += `\n${countStr}`;
    }

    // 학교 폴더명에서 표시 이름 추출 (예: "270. 용인정평초")
    const folderName = state.school_folders.find(f =>
      f.includes(schoolName)
    ) || schoolName;

    Panel.updateSchoolInfo({ school_name: folderName, history_text });
    Panel.setGradeCount(schoolName);

    App.setStepState(1, 'done');
    App.setStepState(2, 'active');
    for (let i = 3; i <= 4; i++) App.setStepState(i, 'idle');

    _el('btn-scan').disabled     = false;
    _el('btn-run').disabled      = true;
    _el('btn-run-diff').disabled = false;

    App.goTab('scan');
    Scan.reset();
  },

  goTab(tab) {
    state.currentTab = tab;
    ['scan', 'run', 'notice', 'diff'].forEach(t => {
      const el = _el(`tab-${t}`);
      if (el) el.classList.toggle('active', t === tab);
    });
    const tabToStep = { scan: 2, run: 3, notice: 4 };
    if (tabToStep[tab] !== undefined) _highlightStep(tabToStep[tab]);
  },

  goStep(idx) {
    const tabMap = { 2: 'scan', 3: 'run', 4: 'notice' };
    if (idx === 0)        App.goBackToSetup();
    else if (idx === 1)   App.resetToSchoolSelect();
    else if (tabMap[idx]) App.goTab(tabMap[idx]);
  },

  setMode(mode) {
    const isMain = mode === 'main';
    _el('btn-mode-main').classList.toggle('active', isMain);
    _el('btn-mode-diff').classList.toggle('active', !isMain);
    App.goTab(isMain ? (state.currentTab === 'diff' ? 'scan' : state.currentTab) : 'diff');
  },

  setStepState(idx, s) {
    const badge = _el(`badge-${idx}`);
    const item  = document.querySelector(`[data-step="${idx}"]`);
    if (!badge || !item) return;
    badge.className = `step-badge${s === 'idle' ? '' : ' ' + s}`;
    item.classList.toggle('active', s === 'active');
  },

  goBackToSetup() {
    if (state.pending_roster_log &&
        !confirm('명단이 아직 기록되지 않았습니다. 설정으로 돌아가시겠습니까?')) return;
    if (!confirm('초기 설정으로 돌아가시겠습니까? 현재 작업 내용이 모두 초기화됩니다.')) return;
    _resetForNewSchool();
    _showPage('setup');
  },

  resetToSchoolSelect() {
    if (state.pending_roster_log &&
        !confirm('명단이 아직 기록되지 않았습니다. 계속하시겠습니까?')) return;
    if (!confirm('새 작업을 시작하시겠습니까? 현재 스캔/실행 결과가 초기화됩니다.')) return;
    _resetForNewSchool();
  },
};

// ──────────────────────────────────────────────
// 초기화 헬퍼
// ──────────────────────────────────────────────
function _resetForNewSchool() {
  state.selected_school    = '';
  state.selected_domain    = '';
  state.current_seq_no     = null;
  state.pending_roster_log = false;
  state.last_scan_logs     = [];
  state.last_run_logs      = [];
  state.last_diff_logs     = [];

  Panel.reset();
  Scan.reset();
  Run.reset();
  Notice.clear();

  App.setStepState(0, 'done');
  for (let i = 1; i <= 4; i++) App.setStepState(i, 'idle');
  App.goTab('scan');

  _el('btn-scan').disabled     = true;
  _el('btn-run').disabled      = true;
  _el('btn-run-diff').disabled = true;
}

function _highlightStep(activeIdx) {
  document.querySelectorAll('.step-item').forEach(el =>
    el.classList.toggle('active', parseInt(el.dataset.step, 10) === activeIdx)
  );
}

function _showPage(page) {
  state.currentPage = page;
  _el('page-setup').style.display = page === 'setup' ? 'flex' : 'none';
  _el('page-main').style.display  = page === 'main'  ? 'flex' : 'none';
}

// ──────────────────────────────────────────────
// 로그 팝업 (alert 없음)
// ──────────────────────────────────────────────
function showLogDialog(title, logs) {
  if (!logs || !logs.length) {
    toast('표시할 로그가 없습니다. 먼저 작업을 실행해 주세요.', 'info');
    return;
  }
  const text = logs
    .filter(l => l.level !== 'debug')
    .map(l => `[${l.level.toUpperCase()}] ${l.message}`)
    .join('\n');

  // <dialog> 기반 로그 뷰어
  let dlg = _el('log-dialog');
  if (!dlg) {
    dlg = document.createElement('dialog');
    dlg.id = 'log-dialog';
    dlg.style.cssText = 'width:640px;max-width:96vw;max-height:80vh;border:none;border-radius:12px;padding:0;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.2)';
    dlg.innerHTML = `
      <div style="padding:16px 20px;border-bottom:1px solid #E5E7EB;font-weight:800;font-size:15px;color:#0F172A" id="log-dlg-title"></div>
      <pre id="log-dlg-body" style="flex:1;overflow-y:auto;padding:16px 20px;font-family:Consolas,'D2Coding',monospace;font-size:12px;line-height:1.6;background:#F8FAFC;white-space:pre-wrap;word-break:break-all;margin:0"></pre>
      <div style="padding:12px 20px;border-top:1px solid #E5E7EB;display:flex;justify-content:flex-end">
        <button style="padding:0 16px;height:34px;background:#2563EB;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer" id="log-dlg-close">닫기</button>
      </div>`;
    document.body.appendChild(dlg);
    _el('log-dlg-close').addEventListener('click', () => dlg.close());
  }
  _el('log-dlg-title').textContent = title;
  _el('log-dlg-body').textContent  = text;
  dlg.showModal();
}

// ──────────────────────────────────────────────
// Notice ctx 생성 (state 직접 참조 격리)
// ──────────────────────────────────────────────
function _noticeCtx() {
  return {
    school_name:       state.selected_school    || '',
    domain:            state.selected_domain    || '',
    school_start_date: state.school_start_date  || '',
  };
}

// ──────────────────────────────────────────────
// 설정 읽기
// ──────────────────────────────────────────────
function _readFullConfig() {
  return {
    work_root:          state.work_root,
    roster_log_path:    state.roster_log_path,
    worker_name:        state.worker_name,
    school_start_date:  state.school_start_date,
    work_date:          state.work_date,
    arrived_date:       state.arrived_date || '',
    last_school:        state.selected_school || '',
    roster_col_map:     state.roster_col_map  || {},
  };
}

// ──────────────────────────────────────────────
// 공통 유틸
// ──────────────────────────────────────────────
function _el(id) { return document.getElementById(id); }

function _todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ──────────────────────────────────────────────
// 개발용 Bridge mock
// ──────────────────────────────────────────────
function _makeMockBridge() {
  const noop = async () => JSON.stringify({ ok: true, data: {} });
  return {
    loadAppConfig:       async () => JSON.stringify({ ok: true, data: { config: {} } }),
    saveAppConfig:       noop,
    inspectWorkRoot:     async () => JSON.stringify({ ok: true, data: { ok: true, errors: [] } }),
    loadSchoolNames:     async () => JSON.stringify({ ok: true, data: { school_names: [] } }),
    getSchoolDomain:     async () => JSON.stringify({ ok: true, data: { domain: '' } }),
    getProjectDirs:      async () => JSON.stringify({ ok: true, data: { dirs: {} } }),
    loadNoticeTemplates: async () => JSON.stringify({ ok: true, data: { templates: {} } }),
    loadWorkHistory:     async () => JSON.stringify({ ok: true, data: { history: {} } }),
    saveWorkHistory:     noop,
    readXlsxMeta:        async () => JSON.stringify({ ok: true, data: { sheets: [], headers: [], rows: [] } }),
    pickWorkFolder:      async () => JSON.stringify({ ok: true, data: { path: '' } }),
    pickRosterLogFile:   async () => JSON.stringify({ ok: true, data: { path: '' } }),
    openFile:            noop,
    openFolder:          noop,
    copyToClipboard:     noop,
    writeWorkResult:     noop,
    writeEmailSent:      noop,
    startScanMain:       async () => JSON.stringify({ ok: true, data: {} }),
    startRunMain:        async () => JSON.stringify({ ok: true, data: {} }),
    startScanDiff:       async () => JSON.stringify({ ok: true, data: {} }),
    startRunDiff:        async () => JSON.stringify({ ok: true, data: {} }),
    startPreview:        async () => JSON.stringify({ ok: true, data: {} }),
    scanFinished:    { connect: () => {} },
    scanFailed:      { connect: () => {} },
    runFinished:     { connect: () => {} },
    runFailed:       { connect: () => {} },
    diffScanFinished:{ connect: () => {} },
    diffScanFailed:  { connect: () => {} },
    diffRunFinished: { connect: () => {} },
    diffRunFailed:   { connect: () => {} },
    previewLoaded:   { connect: () => {} },
    previewFailed:   { connect: () => {} },
  };
}

// ──────────────────────────────────────────────
// 앱 시작
// ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', initApp);
