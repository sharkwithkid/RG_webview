/**
 * scan_tab.js — 스캔 탭 로직
 *
 * 의존: main.js (state, bridge, App, _el, _todayStr)
 *       status_panel.js (Panel)
 * HTML ID: btn-scan, scan-status-badge, scan-message,
 *           scan-tbody (행: data-kind 속성),
 *           spin-{kind}, chk-{kind},
 *           school-kind-warn, school-kind-row, school-kind-select,
 *           btn-toggle-viewer, viewer-body,
 *           preview-file-info, preview-search, preview-warn,
 *           btn-issue-only, btn-dup-only,
 *           preview-table, btn-goto-run
 */

'use strict';

const Scan = (() => {

  // 스캔 완료 후 받은 payload (run_tab에서 참조)
  let _lastScanData = null;

  // 뷰어 상태
  let _viewerOpen   = false;
  let _previewData  = {};       // { kind: { file_name, sheet_name, header_row, data_start_row, columns, rows, issue_rows } }
  let _currentKind  = null;
  let _filterState  = { issue: false, dup: false };
  let _scanWarnState = _makeEmptyWarnState();

  // 구분 → 행 index 고정
  const KIND_ROW = { 신입생: 0, 전입생: 1, 전출생: 2, 교직원: 3 };
  const KIND_KEY = { 신입생: 'freshmen', 전입생: 'transfer_in', 전출생: 'transfer_out', 교직원: 'teachers' };

  // ──────────────────────────────────────────────
  // 스캔 시작
  // ──────────────────────────────────────────────
  async function start() {
    if (state.isScanning) return;
    if (!state.work_root)        { _setMessage('작업 폴더가 설정되지 않았습니다.'); return; }
    if (!state.selected_school)  { _setMessage('학교를 먼저 선택해 주세요.'); return; }
    if (!state.school_start_date){ _setMessage('개학일이 설정되지 않았습니다.'); return; }
    if (!state.work_date)        { _setMessage('작업일이 설정되지 않았습니다.'); return; }

    state.isScanning = true;
    _el('btn-scan').disabled = true;
    _setBadge('running', '스캔 중');
    _setMessage('스캔 중...');

    _previewData = {};
    _currentKind = null;
    _lastScanData = null;
    _el('btn-run').disabled = true;
    App.setFloatingNext(false, null);
    if (typeof Panel !== 'undefined' && Panel.updateRosterMapBtn) Panel.updateRosterMapBtn(null);
    _hideSchoolKindWarn();
    _scanWarnState = _makeEmptyWarnState();
    _hideScanWarnCard();
    _el('preview-warn').textContent = '';
    _el('preview-file-info').textContent = '';
    _setMessage('스캔 중...');
    _updateGradeMap({ need_roster: false });

    const params = {
      work_root:          state.work_root,
      school_name:        state.selected_school,
      school_start_date:  state.school_start_date,
      work_date:          state.work_date,
      roster_xlsx:        state.roster_log_path || '',
      col_map:            state.roster_col_map  || {},
    };

    const res = JSON.parse(await bridge.startScanMain(JSON.stringify(params)));
    if (!res.ok) {
      // 동기 검증 실패 (파라미터 오류 등)
      state.isScanning = false;
      _el('btn-scan').disabled = false;
      _setBadge('err', '스캔 실패');
      _setMessage(res.error || '스캔 시작 실패');
    }
    // 비동기 완료는 main.js → bridge.scanFinished → onFinished / onFailed
  }

  // ──────────────────────────────────────────────
  // 스캔 완료 (main.js bridge.scanFinished에서 호출)
  // ──────────────────────────────────────────────
  function onFinished(data) {
    _el('btn-scan').disabled = false;
    _lastScanData = data;

    if (!data.ok) {
      const errLog = (data.logs || []).find(l => l.level === 'error');
      const errMsg = errLog ? errLog.message : '스캔 중 오류가 발생했습니다.';
      _setBadge('err', '스캔 실패');
      _setMessage(errMsg);
      _hideSchoolKindWarn();
      App.setStepState(2, 'warn');
      toast(errMsg, 'err', 6000);
      return;
    }

    // 스텝 상태
    App.setStepState(2, 'done');

    // 학교 구분 판별 실패 여부
    const kindWarn = (data.logs || []).some(l =>
      l.level === 'warn' && l.message.includes('학교 구분을 자동으로 판별하지 못했습니다')
    );
    _setSchoolKindWarn(kindWarn);

    // 경고 상태 정규화
    const errLogs  = (data.logs || []).filter(l => l.level === 'error');
    _scanWarnState = _buildWarnState(data);

    if (errLogs.length) {
      _setBadge('err', '오류');
      _setMessage('');
      _showScanWarnCard(errLogs.map(l => String(l.message || '').trim()).filter(Boolean), 'error');
    } else if (_scanWarnState.hasWarn) {
      _setBadge('warn', '경고');
      _setMessage('');
      _showScanWarnCard(_scanWarnState.messages, 'warn');
    } else {
      _setBadge('ok', '스캔 완료');
      _setMessage('');
      _hideScanWarnCard();
    }
    _applyScanRowWarnStyles();

    // 학년도 아이디 규칙 갱신 + 명부 버튼 상태
    _updateGradeMap(data);
    if (typeof Panel !== 'undefined' && Panel.updateRosterMapBtn) Panel.updateRosterMapBtn(data.roster_path || null);

    const hasTransfer = (data.items || []).some(i => i.kind === '전입생');
    const hasWithdraw = (data.items || []).some(i => i.kind === '전출생');
    if ((hasTransfer || hasWithdraw) && !data.roster_path) {
      const previewWarn = _el('preview-warn');
      if (previewWarn) {
        previewWarn.textContent = '학생명부가 필요합니다. 명부를 추가해 주세요.';
        previewWarn.style.color = '#B91C1C';
      }
    }

    // 스캔 표 갱신
    _applyScanTable(data.items || []);

    // 뷰어 자동 펼침 + 첫 행 미리보기 요청
    _openViewer();
    _loadFirstPreview(data.items || []);

    // 실행 버튼 활성 + 플로팅 버튼
    if (data.can_execute) {
      _el('btn-run').disabled = false;
      App.setFloatingNext(true, 'run');
    } else {
      _el('btn-run').disabled = true;
      App.setFloatingNext(false, null);
    }

    // 명부 기준일 불일치 처리
    if (data.roster_date_mismatch && data.roster_basis_date) {
      _handleRosterDateMismatch(data.roster_basis_date);
    }
  }

  function onFailed(error) {
    _el('btn-scan').disabled = false;
    _setBadge('err', '스캔 실패');
    _setMessage('예기치 못한 오류가 발생했습니다.');
    App.setStepState(2, 'warn');
    toast('스캔 오류 — 스캔 로그 보기에서 자세한 내용을 확인하세요.', 'err');
  }

  // ──────────────────────────────────────────────
  // 스캔 표 갱신
  // ──────────────────────────────────────────────
  function _applyScanTable(items) {
    const presentKinds = new Set(items.map(i => i.kind));

    items.forEach(item => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${item.kind}"]`);
      if (!tr) return;

      const cells = tr.querySelectorAll('td');
      const kindWarn = !!_scanWarnState.byKind[item.kind]?.messages?.length;

      tr.classList.toggle('scan-row-warn', kindWarn);

      if (cells[1]) {
        cells[1].className = kindWarn ? 'file-link file-link-warn' : 'file-link';
        cells[1].textContent = item.file_name || '';
        cells[1].title = kindWarn
          ? `경고 있음 · 클릭: 뷰어로 보기 · 더블클릭: 파일 열기`
          : '클릭: 뷰어로 보기 · 더블클릭: 파일 열기';
        cells[1].onclick = () => {
          document.querySelectorAll('#scan-tbody tr.viewer-active')
            .forEach(r => r.classList.remove('viewer-active'));
          tr.classList.add('viewer-active');
          _requestPreview(item.kind);
        };
        cells[1].ondblclick = () => { if (item.file_path) bridge.openFile(item.file_path); };
      }
      if (cells[2]) cells[2].textContent = item.sheet_name || '';

      const spinVal = _el(`spin-${item.kind}`);
      if (spinVal && item.data_start_row != null) {
        spinVal.textContent = String(item.data_start_row);
        spinVal.style.color = '#0F172A';
      }
    });

    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;

      if (!presentKinds.has(kind)) {
        const cells = tr.querySelectorAll('td');
        tr.classList.remove('scan-row-warn', 'viewer-active');
        if (cells[1]) { cells[1].className = ''; cells[1].textContent = ''; cells[1].onclick = null; cells[1].ondblclick = null; }
        if (cells[2]) cells[2].textContent = '';
        const spinVal = _el(`spin-${kind}`);
        if (spinVal) { spinVal.textContent = '-'; spinVal.style.color = '#94A3B8'; }
        const chk = _el(`chk-${kind}`);
        if (chk) { chk.checked = false; chk.disabled = true; }
      } else {
        const chk = _el(`chk-${kind}`);
        if (chk) chk.disabled = false;
      }
    });
  }

  // ──────────────────────────────────────────────
  // 스핀 (수정 시작행 +/-)
  // ──────────────────────────────────────────────
  function spin(kind, delta) {
    const el = _el(`spin-${kind}`);
    if (!el) return;
    const cur = parseInt(el.textContent, 10);
    if (isNaN(cur)) {
      if (delta > 0) { el.textContent = '1'; el.style.color = '#0F172A'; }
      return;
    }
    const next = cur + delta;
    if (next < 1) return;
    el.textContent = String(next);
    el.style.color = '#0F172A';
  }

  // 수정 시작행 오버라이드 수집 (startRunMain 호출 시 사용)
  function getLayoutOverrides() {
    const overrides = {};
    Object.entries(KIND_KEY).forEach(([kind, key]) => {
      const el = _el(`spin-${kind}`);
      if (!el) return;
      const v = parseInt(el.textContent, 10);
      if (v > 0) overrides[key] = { data_start_row: v };
    });
    // 학년도 아이디 규칙 오버라이드
    const gradeYears = Panel.getGradeOverrides();
    if (Object.keys(gradeYears).length) overrides.grade_year_map = gradeYears;
    return overrides;
  }

  // 학교 구분 오버라이드 (학교 구분 자동 판별 실패 시)
  function getSchoolKindOverride() {
    const row = _el('school-kind-row');
    if (!row || row.style.display === 'none') return null;
    return _el('school-kind-select')?.value || null;
  }

  // ──────────────────────────────────────────────
  // 미리보기 요청 (startPreview → bridge)
  // ──────────────────────────────────────────────
  async function _requestPreview(kind) {
    if (_previewData[kind]) {
      _currentKind = kind;
      _renderPreview(kind);
      return;
    }
    if (state.isPreviewLoading) {
      _currentKind = kind;
      _el('preview-warn').textContent = `${kind} 미리보기 로딩 중...`;
      return;
    }

    // 스캔 데이터에서 파일 메타 찾기
    const item = (_lastScanData?.items || []).find(i => i.kind === kind);
    if (!item || !item.file_path) {
      _el('preview-warn').textContent = `${kind} 파일 경로 정보가 없습니다.`;
      return;
    }

    _currentKind = kind;
    state.isPreviewLoading = true;
    _el('preview-warn').textContent = `${kind} 미리보기 로딩 중...`;

    const params = {
      kind,
      file_path:      item.file_path,
      sheet_name:     item.sheet_name     || '',
      header_row:     item.header_row     || 1,
      data_start_row: item.data_start_row || 2,
    };

    const res = JSON.parse(await bridge.startPreview(JSON.stringify(params)));
    if (!res.ok) {
      state.isPreviewLoading = false;
      _el('preview-warn').textContent = res.error || '미리보기 시작 실패';
    }
    // 완료는 main.js → bridge.previewLoaded → onPreviewLoaded
  }

  function onPreviewLoaded(payload) {
    // 열 매핑 다이얼로그용 미리보기 분기
    if (payload.kind === 'col_map') {
      ColMap.onPreviewLoaded(payload);
      return;
    }

    const {
      kind, columns, rows, total_count, truncated,
      source_file, sheet_name,
      header_row, data_start_row,
      issue_rows,
    } = payload;

    _previewData[kind] = {
      columns, rows, total_count, truncated,
      source_file, sheet_name,
      header_row:     header_row     ?? null,
      data_start_row: data_start_row ?? null,
      issue_rows:     issue_rows     || [],
    };
    if (_currentKind && _currentKind !== kind && !_previewData[_currentKind]) {
      // 다른 kind로 전환됐지만 캐시 없음 → 렌더 없이 대기
      // (다음 클릭 시 재요청됨 — 재귀 호출 금지)
      _el('preview-warn').textContent = `${_currentKind} 파일을 클릭해서 미리보기를 열어 주세요.`;
    } else {
      _currentKind = _currentKind || kind;
      _renderPreview(_currentKind);
    }
  }

  function onPreviewFailed(kind, error) {
    _el('preview-warn').textContent = `${kind} 파일을 미리볼 수 없습니다.`;
  }

  // ──────────────────────────────────────────────
  // 첫 항목 자동 미리보기
  // ──────────────────────────────────────────────
  function _loadFirstPreview(items) {
    const order = ['신입생', '전입생', '전출생', '교직원'];
    const first = order.find(k => items.some(i => i.kind === k));
    if (first) _requestPreview(first);
  }

  // ──────────────────────────────────────────────
  // 미리보기 렌더링
  // ──────────────────────────────────────────────
  function _renderPreview(kind) {
    const data = _previewData[kind];
    if (!data) return;

    const headerInfo = data.header_row     != null ? ` | 헤더행: ${data.header_row}`     : '';
    const startInfo  = data.data_start_row != null ? ` | 시작행: ${data.data_start_row}` : '';
    _el('preview-file-info').textContent =
      `파일: ${data.source_file || '-'} | 시트: ${data.sheet_name || '-'}` +
      headerInfo + startInfo +
      (data.truncated ? ` | ${data.rows.length}행까지 표시` : '');

    const previewWarn = _el('preview-warn');
    if (previewWarn) {
      previewWarn.textContent = '';
      previewWarn.style.color = '';
    }

    const meta = _renderTable(data);
    _syncPreviewWarn(kind, meta);
  }

function _renderTable(data) {
  const keyword   = (_el('preview-search')?.value || '').trim().toLowerCase();
  const issueOnly = _filterState.issue;
  const dupOnly   = _filterState.dup;

  const columns  = data.columns || [];
  const rows     = data.rows    || [];
  const issueSet = new Set(data.issue_rows || []);

  const nameCol  = columns.findIndex(h => ['성명','이름','학생이름'].some(k => h.includes(k)));
  const gradeCol = columns.findIndex(h => h.includes('학년'));
  const classCol = columns.findIndex(h => ['반','학급'].some(k => h.includes(k)) && !h.includes('학년'));

  const dupSet = new Set();
  if (nameCol >= 0) {
    const cnt = {};
    rows.forEach((r, i) => {
      const nm    = (r[nameCol] || '').replace(/[A-Z]+$/, '').trim();
      const grade = gradeCol >= 0 ? (r[gradeCol] || '') : '';
      const key   = `${grade}||${nm}`;
      if (nm) cnt[key] = (cnt[key] || []).concat(i);
    });
    Object.values(cnt).forEach(idxs => { if (idxs.length >= 2) idxs.forEach(i => dupSet.add(i)); });
  }

  const rowIssueMap = new Map();
  const addIssue = (i, msg) => {
    if (!msg) return;
    const arr = rowIssueMap.get(i) || [];
    if (!arr.includes(msg)) arr.push(msg);
    rowIssueMap.set(i, arr);
  };

  rows.forEach((r, i) => {
    const vals = r.map(v => String(v || '').trim());
    const isBlankRow = vals.every(v => !v);

    if (isBlankRow) addIssue(i, '빈 행 확인');
    else if (issueSet.has(i)) addIssue(i, '필수값 확인');

    if (nameCol >= 0) {
      const nm = String(r[nameCol] || '').trim();
      if (!nm && !isBlankRow) addIssue(i, '이름 확인');
      if (nm) {
        const hasKo  = /[가-힣]/.test(nm);
        const hasEn  = /[A-Za-z]/.test(nm);
        const hasNum = /\d/.test(nm);
        const hasSpc = /[^\w가-힣\s]/.test(nm);
        if (hasNum || (hasKo && hasEn) || (hasSpc && (hasKo || hasEn))) addIssue(i, '이름 확인');
      }
    }
    if (gradeCol >= 0) {
      const gv = String(r[gradeCol] || '').trim();
      if (!gv && !isBlankRow) addIssue(i, '학년 확인');
      else if (gv && !/^\d+$/.test(gv)) addIssue(i, '학년 확인');
    }
    if (classCol >= 0) {
      const cv = String(r[classCol] || '').trim();
      if (!cv && !isBlankRow) addIssue(i, '반 확인');
      else if (cv && !/^\d+$/.test(cv)) addIssue(i, '반 확인');
    }
  });

  const filtered = rows.reduce((acc, row, i) => {
    const rowText = row.join(' ').toLowerCase();
    if (keyword && !rowText.includes(keyword)) return acc;
    const issues = rowIssueMap.get(i) || [];
    const isIssue = issues.length > 0;
    const isDup = dupSet.has(i);
    if (issueOnly && !isIssue) return acc;
    if (dupOnly && !isDup) return acc;
    acc.push({ row, i, issues, isIssue, isDup });
    return acc;
  }, []);

  const table    = _el('preview-table');
  const thead    = table.querySelector('thead');
  const tbody    = table.querySelector('tbody');
  const startRow = data.data_start_row ?? 1;

  const headers = [...columns, '비고'];
  thead.innerHTML = '<tr><th style="width:40px;color:#94A3B8;font-weight:600;text-align:center">#</th>' +
    headers.map(h => `<th>${_esc(h)}</th>`).join('') + '</tr>';
  tbody.innerHTML = filtered.map(({ row, i, issues, isIssue, isDup }) => {
    const excelRow = startRow + i;
    const finalCls = isIssue ? 'row-warn' : isDup ? 'row-dup' : '';
    const cells = row.map(v => `<td>${_esc(v)}</td>`).join('') +
      `<td>${issues.length ? _esc(issues.join(', ')) : ''}</td>`;
    return `<tr class="${finalCls}"><td style="color:#94A3B8;font-size:11px;text-align:center;user-select:none">${excelRow}</td>${cells}</tr>`;
  }).join('');

  const kindEntry = _scanWarnState.byKind[_currentKind] || { messages: [], issueRows: new Set(), suspectCount: 0, rowIssues: new Map() };
  kindEntry.rowIssues = rowIssueMap;
  kindEntry.issueRows = new Set([...rowIssueMap.keys()]);
  kindEntry.suspectCount = rowIssueMap.size;
  _scanWarnState.byKind[_currentKind] = kindEntry;

  return {
    issueCount: rowIssueMap.size,
    rowIssueMap,
    firstIssueExcelRow: filtered.find(x => x.isIssue)?.i != null
      ? startRow + filtered.find(x => x.isIssue).i
      : null,
  };
}

function _syncPreviewWarn(kind, meta = {}) {
  const previewWarn = _el('preview-warn');
  if (!previewWarn) return;

  const kindState = _scanWarnState.byKind[kind] || { messages: [], issueRows: new Set(), suspectCount: 0, rowIssues: new Map() };
  const rowIssueMap = meta.rowIssueMap || kindState.rowIssues || new Map();
  const count = typeof meta.issueCount === 'number' ? meta.issueCount : rowIssueMap.size;

  if (!count) {
    previewWarn.textContent = '';
    previewWarn.style.color = '';
    _refreshWarnUI();
    return;
  }

  const firstEntry = rowIssueMap.entries().next().value;
  const firstIdx = firstEntry?.[0];
  const firstReasons = firstEntry?.[1] || [];
  const excelRow = (_previewData[kind]?.data_start_row ?? 1) + (typeof firstIdx === 'number' ? firstIdx : 0);
  const lead = `${excelRow}행 ${firstReasons[0] || '확인'}`;
  const suffix = count > 1 ? ` 외 ${count - 1}건` : '';

  previewWarn.innerHTML = `<div class="pw-inline">⚠ 경고 · ${_esc(lead)}${_esc(suffix)}</div>`;
  previewWarn.style.color = '#92400E';

  _refreshWarnUI();
}

function _refreshWarnUI() {
  const errLogs = (state.last_scan_logs || []).filter(l => l.level === 'error');
  if (errLogs.length) {
    _setBadge('err', '오류');
    _showScanWarnCard(errLogs.map(l => String(l.message || '').trim()).filter(Boolean), 'error');
    _applyScanRowWarnStyles();
    return;
  }

  _scanWarnState.hasWarn = (_scanWarnState.messages || []).length > 0 || Object.values(_scanWarnState.byKind || {}).some(v =>
    (v.messages && v.messages.length) || (v.issueRows && v.issueRows.size) || (v.suspectCount || 0) > 0
  );

  if (_scanWarnState.hasWarn) {
    _setBadge('warn', '경고');
    _showScanWarnCard(_scanWarnState.messages, 'warn');
  } else {
    _setBadge('ok', '스캔 완료');
    _hideScanWarnCard();
  }
  _applyScanRowWarnStyles();
}

  function _applyScanRowWarnStyles() {
    const errorKinds = new Set(
      (state.last_scan_logs || [])
        .filter(l => l.level === 'error')
        .map(l => _kindFromWarnMessage(String(l.message || '').trim()))
        .filter(Boolean)
    );

    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      const hasKindWarn = !!(_scanWarnState.byKind[kind]?.messages?.length || _scanWarnState.byKind[kind]?.issueRows?.size || _scanWarnState.byKind[kind]?.suspectCount);
      const hasKindError = errorKinds.has(kind);
      tr.classList.toggle('scan-row-warn', hasKindWarn && !hasKindError);
      tr.classList.toggle('scan-row-error', hasKindError);
      const fileCell = tr.querySelectorAll('td')[1];
      if (fileCell && fileCell.textContent) {
        fileCell.classList.toggle('file-link-warn', hasKindWarn && !hasKindError);
        fileCell.classList.toggle('file-link-error', hasKindError);
      }
    });
  }

  function _makeEmptyWarnState() {
    return {
      hasWarn: false,
      messages: [],
      byKind: {},
    };
  }

  function _buildWarnState(data) {
    const stateObj = _makeEmptyWarnState();
    const items = data.items || [];
    items.forEach(item => {
      const kind = item.kind;
      const entry = stateObj.byKind[kind] || { messages: [], issueRows: new Set(), suspectCount: 0, rowIssues: new Map() };
      const warningText = String(item.warning || '').trim();
      if (warningText) {
        warningText.split(/\n+/).map(s => s.trim()).filter(Boolean).forEach(msg => _pushWarnMessage(stateObj, kind, msg));
      }
      (item.issue_rows || []).forEach(r => entry.issueRows.add(r));
      stateObj.byKind[kind] = entry;
    });

    (data.logs || [])
      .filter(l => l.level === 'warn')
      .forEach(l => {
        const msg = String(l.message || '').trim();
        const kind = _kindFromWarnMessage(msg);
        _pushWarnMessage(stateObj, kind, msg);
      });

    stateObj.hasWarn = stateObj.messages.length > 0 || Object.values(stateObj.byKind).some(v => (v.issueRows && v.issueRows.size) || v.suspectCount);
    return stateObj;
  }

  function _pushWarnMessage(stateObj, kind, msg) {
    if (!msg) return;
    if (!stateObj.messages.includes(msg)) stateObj.messages.push(msg);
    if (kind) {
      const entry = stateObj.byKind[kind] || { messages: [], issueRows: new Set(), suspectCount: 0, rowIssues: new Map() };
      if (!entry.messages.includes(msg)) entry.messages.push(msg);
      stateObj.byKind[kind] = entry;
    }
  }

  function _kindFromWarnMessage(msg) {
    if (!msg) return null;
    if (msg.includes('신입생')) return '신입생';
    if (msg.includes('전입생')) return '전입생';
    if (msg.includes('전출생')) return '전출생';
    if (msg.includes('교직원') || msg.includes('교사')) return '교직원';
    return null;
  }

function _normalizeScanCardMessage(msg, mode) {
  const s = String(msg || '').trim();
  if (!s) return '';
  if (/헤더를 찾을 수 없습니다|구조를 읽을 수 없습니다/.test(s)) {
    return '헤더를 찾을 수 없습니다. 파일 구조를 확인해 주세요.';
  }
  if (/시트가\s*\d+개 있습니다/.test(s) || /시트가 여러 개 있습니다/.test(s)) {
    return '시트가 여러 개 있습니다. 첫 번째 시트만 사용합니다.';
  }
  return s.replace(/^\[(WARN|ERROR)\]\s*/, '');
}

function _showScanWarnCard(messages, mode = 'warn') {
  const el = _el('scan-warn-card');
  if (!el) return;
  const lines = Array.from(new Set((Array.isArray(messages) ? messages : []).map(msg => _normalizeScanCardMessage(msg, mode)).filter(Boolean)));
  if (!lines.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    el.classList.remove('error');
    return;
  }
  el.classList.toggle('error', mode === 'error');
  const head = mode === 'error'
    ? `<div class="warn-line">오류 ${lines.length}건이 있습니다. 파일 구조를 확인해 주세요.</div>`
    : `<div class="warn-line">경고 ${lines.length}건이 있습니다. 선택 파일에서 확인해 주세요.</div>`;
  const body = lines.slice(0, 3).map(msg => `<div class="warn-line">• ${_esc(msg)}</div>`).join('');
  const tail = lines.length > 3 ? `<div class="warn-line">외 ${lines.length - 3}건</div>` : '';
  el.innerHTML = head + body + tail;
  el.style.display = 'block';
}

  function _hideScanWarnCard() {
    const el = _el('scan-warn-card');
    if (!el) return;
    el.style.display = 'none';
    el.innerHTML = '';
    el.classList.remove('error');
  }

  // ──────────────────────────────────────────────
  // 필터 토글
  // ──────────────────────────────────────────────
function toggleFilter(key) {
  _filterState[key] = !_filterState[key];
  const btnId = { issue: 'btn-issue-only', dup: 'btn-dup-only' }[key];
  _el(btnId)?.classList.toggle('active', _filterState[key]);
  if (_currentKind) _renderPreview(_currentKind);
}

  function filterPreview() {
    if (_currentKind) _renderPreview(_currentKind);
  }

  // ──────────────────────────────────────────────
  // 뷰어 펼치기/접기
  // ──────────────────────────────────────────────
  function toggleViewer() {
    _viewerOpen = !_viewerOpen;
    _el('viewer-body').style.display = _viewerOpen ? '' : 'none';
    _el('btn-toggle-viewer').textContent = _viewerOpen ? '접기 ▴' : '펼치기 ▾';
  }

  function _openViewer() {
    if (_viewerOpen) return;
    _viewerOpen = true;
    _el('viewer-body').style.display = '';
    _el('btn-toggle-viewer').textContent = '접기 ▴';
  }

  // ──────────────────────────────────────────────
  // 학교 구분 경고 UI
  // ──────────────────────────────────────────────
  function _setSchoolKindWarn(show) {
    _el('school-kind-warn').style.display = show ? '' : 'none';
    const row = _el('school-kind-row');
    if (row) row.style.display = show ? 'flex' : 'none';
  }

  function _hideSchoolKindWarn() { _setSchoolKindWarn(false); }

  // ──────────────────────────────────────────────
  // 학년도 아이디 규칙 갱신 (StatusPanel)
  // ──────────────────────────────────────────────
  function _updateGradeMap(data) {
    if (!data.need_roster) {
      Panel.updateGradeMap('not_needed');
      return;
    }
    const gym = data.grade_year_map;
    if (gym && Object.keys(gym).length) {
      Panel.updateGradeMap('ok', gym);
    } else {
      Panel.updateGradeMap('no_roster');
    }
  }

  // ──────────────────────────────────────────────
  // 명부 기준일 불일치 처리 (커스텀 모달)
  // ──────────────────────────────────────────────
  function _handleRosterDateMismatch(basisDate) {
    const workDate = state.work_date;

    // 배경 + 모달 생성
    const backdrop = document.createElement('div');
    backdrop.className = 'confirm-modal-backdrop';

    backdrop.innerHTML = `
      <div class="confirm-modal">
        <div class="confirm-modal-title">명부 기준일 설정</div>
        <div class="confirm-modal-body">
          학생명부 마지막 수정일과 작업일이 다릅니다.<br>
          어느 날짜를 명부 기준일로 사용할까요?
        </div>
        <div class="confirm-modal-options">
          <label class="confirm-modal-option selected" id="cm-opt-basis">
            <input type="radio" name="cm-date" value="basis" checked>
            <div>
              <div class="confirm-modal-option-label">수정일 사용 — ${basisDate}</div>
              <div class="confirm-modal-option-desc">명부 파일의 마지막 수정일을 기준으로 합니다.</div>
            </div>
          </label>
          <label class="confirm-modal-option" id="cm-opt-work">
            <input type="radio" name="cm-date" value="work">
            <div>
              <div class="confirm-modal-option-label">작업일로 재스캔 — ${workDate}</div>
              <div class="confirm-modal-option-desc">오늘 작업일을 기준으로 다시 스캔합니다.</div>
            </div>
          </label>
        </div>
        <div class="confirm-modal-footer">
          <button class="btn-primary" id="cm-date-confirm" style="height:36px;padding:0 20px">확인</button>
        </div>
      </div>`;

    document.body.appendChild(backdrop);

    // 라디오 선택 시 스타일 갱신
    backdrop.querySelectorAll('input[type="radio"]').forEach(radio => {
      radio.addEventListener('change', () => {
        backdrop.querySelectorAll('.confirm-modal-option').forEach(opt =>
          opt.classList.remove('selected')
        );
        radio.closest('.confirm-modal-option').classList.add('selected');
      });
    });

    // 확인 버튼
    backdrop.querySelector('#cm-date-confirm').addEventListener('click', () => {
      const selected = backdrop.querySelector('input[name="cm-date"]:checked')?.value;
      backdrop.remove();
      if (selected === 'work') {
        _rescanWithBasisDate(workDate);
      }
      // 'basis' 선택 시 그냥 진행 (아무것도 안 함)
    });
  }

  async function _rescanWithBasisDate(basisDate) {
    if (state.isScanning) return;
    state.isScanning = true;
    _el('btn-scan').disabled = true;
    _setBadge('running', '재스캔 중');

    // 이전 스캔 결과 초기화
    _previewData = {};
    _currentKind = null;
    _lastScanData = null;
    _el('btn-run').disabled = true;
    App.setFloatingNext(false, null);
    if (typeof Panel !== 'undefined' && Panel.updateRosterMapBtn) Panel.updateRosterMapBtn(null);
    _hideSchoolKindWarn();
    _scanWarnState = _makeEmptyWarnState();
    _hideScanWarnCard();
    _el('preview-warn').textContent = '';
    _el('preview-file-info').textContent = '';
    _setMessage('재스캔 중...');
    _updateGradeMap({ need_roster: false });

    const params = {
      work_root:          state.work_root,
      school_name:        state.selected_school,
      school_start_date:  state.school_start_date,
      work_date:          state.work_date,
      roster_xlsx:        state.roster_log_path || '',
      col_map:            state.roster_col_map  || {},
      roster_basis_date:  basisDate,
    };

    const res = JSON.parse(await bridge.startScanMain(JSON.stringify(params)));
    if (!res.ok) {
      state.isScanning = false;
      _el('btn-scan').disabled = false;
      _setBadge('err', '스캔 실패');
      _setMessage(res.error || '재스캔 시작 실패');
    }
  }

  // ──────────────────────────────────────────────
  // 로그 팝업
  // ──────────────────────────────────────────────
  function showLog() {
    showLogDialog('스캔 로그', state.last_scan_logs);
  }

  function applyManualGradeReady(overrides) {
    if (!_lastScanData || !_lastScanData.can_execute_after_input) return false;

    const hasStrictRosterError = (state.last_scan_logs || []).some(l =>
      l.level === 'error' && String(l.message || '').includes('학생명부가 필요합니다.')
    );
    if (hasStrictRosterError) return false;

    _lastScanData.grade_year_map = {
      ...(_lastScanData.grade_year_map || {}),
      ...(overrides || {}),
    };
    _lastScanData.can_execute = true;

    _setBadge(_scanWarnState.hasWarn ? 'warn' : 'ok', _scanWarnState.hasWarn ? '경고' : '스캔 완료');
    _setMessage('학년도 아이디 규칙이 적용되었습니다. 바로 실행할 수 있습니다.');
    _el('btn-run').disabled = false;
    App.setFloatingNext(true, 'run');

    const previewWarn = _el('preview-warn');
    if (previewWarn && previewWarn.textContent.includes('학년도 아이디 규칙')) {
      previewWarn.textContent = '';
    }

    return true;
  }

  // ──────────────────────────────────────────────
  // 초기화 (학교 변경 시)
  // ──────────────────────────────────────────────
  function reset() {
    _lastScanData  = null;
    _previewData   = {};
    _currentKind   = null;
    if (typeof Panel !== 'undefined' && Panel.updateRosterMapBtn) Panel.updateRosterMapBtn(null);
    _filterState   = { issue: false, dup: false };

    _setBadge('idle', '스캔 전');
    _setMessage('');
    _hideSchoolKindWarn();
    _scanWarnState = _makeEmptyWarnState();
    _hideScanWarnCard();

    // 스캔 표 초기화
    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      const cells = tr.querySelectorAll('td');
      [1, 2].forEach(c => { if (cells[c]) cells[c].textContent = ''; });
      const spinVal = _el(`spin-${kind}`);
      if (spinVal) { spinVal.textContent = '-'; spinVal.style.color = '#94A3B8'; }
      const chk = _el(`chk-${kind}`);
      if (chk) { chk.checked = false; chk.disabled = false; }
    });

    // 뷰어
    if (_viewerOpen) toggleViewer();
    _el('preview-file-info').textContent = '파일: - | 시트: - | 헤더행: - | 시작행: -';
    _el('preview-warn').textContent      = '학교를 선택하고 스캔을 실행해 주세요.';
    const table = _el('preview-table');
    if (table) { table.querySelector('thead').innerHTML = ''; table.querySelector('tbody').innerHTML = ''; }

    _el('btn-run').disabled = true;
    App.setFloatingNext(false, null);

    // 필터 버튼 초기화
    ['btn-issue-only','btn-dup-only'].forEach(id => _el(id)?.classList.remove('active'));
  }

  // 실행 탭으로 이동 (확인 체크박스 검증)
  function goToRun() {
    const lastData = _lastScanData;
    if (!lastData) return;

    const presentKinds = (lastData.items || []).map(i => i.kind);
    const unchecked = presentKinds.filter(kind => {
      const chk = _el(`chk-${kind}`);
      return chk && !chk.checked;
    });

    if (unchecked.length) {
      toast(`${unchecked.join(', ')} 파일의 시작행을 확인해 주세요`, 'warn', 4000);
      return;
    }

    App.goTab('run');
  }

  // 외부(run_tab)에서 마지막 스캔 데이터 참조용
  function getLastScanData() { return _lastScanData; }

  // ──────────────────────────────────────────────
  // 내부 헬퍼
  // ──────────────────────────────────────────────
  function _setBadge(type, text) {
    const el = _el('scan-status-badge');
    if (!el) return;
    el.className  = `status-badge badge-${type}`;
    el.textContent = text;
  }

  function _setMessage(msg) {
    const el = _el('scan-message');
    if (el) el.textContent = msg;
  }

  function _esc(str) {
    return String(str ?? '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  // ──────────────────────────────────────────────
  // Public
  // ──────────────────────────────────────────────
  return {
    start, onFinished, onFailed,
    onPreviewLoaded, onPreviewFailed,
    spin, toggleFilter, filterPreview,
    toggleViewer, showLog, reset, goToRun,
    getLastScanData, getLayoutOverrides, getSchoolKindOverride,
    applyManualGradeReady,
  };

})();
