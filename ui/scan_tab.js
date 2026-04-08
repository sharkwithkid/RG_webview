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
    _hideScanWarnCard();
    _el('preview-warn').textContent = '';
    _el('preview-file-info').textContent = '';
    _setMessage('스캔 중...');
    _updateGradeMap({ need_roster: false });

    const params = {
      work_root:            state.work_root,
      school_name:          state.selected_school,
      school_start_date:    state.school_start_date,
      work_date:            state.work_date,
      roster_xlsx:          state.roster_log_path || '',
      col_map:              state.roster_col_map  || {},
      school_kind_override: state.school_kind_override || null,
    };

    const res = JSON.parse(await bridge.startScanMain(JSON.stringify(params)));
    if (!res.ok) {
      // 동기 검증 실패 (파라미터 오류 등)
      state.isScanning = false;
      _el('btn-scan').disabled = false;
      _setBadge('err', '실패');
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
      const status = data.status || null;
      const events = data.events || [];

      // 학교 구분 판별 불가 → 모달로 선택 유도 (스캔 재시작 필요)
      if (data.school_kind_needs_choice) {
        state.isScanning = false;
        _el('btn-scan').disabled = false;
        _showSchoolKindModal();
        return;
      }

      _setBadge('err', '실패');
      _setMessage('');
      _setSchoolKindWarn(false);
      App.setStepState(2, 'warn');
      const errMsgsAll = events.filter(e => e.level === 'error').map(e => e.message);
      if (errMsgsAll.length) {
        _showScanWarnCard(errMsgsAll, 'error', status);
      } else {
        const statusErrs = UICommon.getStatusMessages(status, ['error']);
        if (statusErrs.length) _showScanWarnCard(statusErrs, 'error', status);
        else _showScanWarnCard(['예기치 못한 오류가 발생했습니다. 스캔 로그를 확인해 주세요.'], 'error', status);
      }
      return;
    }

    // 스텝 상태
    App.setStepState(2, 'done');

    // 학교 구분 선택 필요 여부 (ok=True로 오는 경우는 현재 없지만 안전 처리)
    if (data.school_kind_needs_choice) { state.isScanning = false; _el('btn-scan').disabled = false; _showSchoolKindModal(); return; }
    _setSchoolKindWarn(false);

    // 뱃지 + 카드 — events 직접 사용
    const status = data.status || null;
    const events = data.events || [];
    StatusUI.renderBadge('scan-status-badge', status?.badge, '완료');
    _setMessage('');
    const errMsgs  = events.filter(e => e.level === 'error').map(e => e.message);
    const holdMsgs = events.filter(e => e.level === 'hold').map(e => e.message);
    const warnMsgs = events.filter(e => e.level === 'warn').map(e => e.message);
    // warn + hold 함께 있으면 합쳐서 표시 (경고 먼저, 보류 뒤)
    const combinedMsgs = [...warnMsgs, ...holdMsgs];
    if (errMsgs.length)           _showScanWarnCard(errMsgs,        'error', status);
    else if (combinedMsgs.length) _showScanWarnCard(combinedMsgs,   'warn',  null);
    else                          _hideScanWarnCard();

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
    _applyRowMarkStyles(data.row_marks || []);

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
    _setBadge('err', '실패');
    _setMessage('예기치 못한 오류가 발생했습니다. 스캔 로그를 확인해 주세요.');
    App.setStepState(2, 'warn');
    // 토스트 제거 — _setMessage에 내용 표시
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
      const kindWarn  = item.severity === 'warn';
      const kindError = item.severity === 'error';

      tr.classList.toggle('scan-row-warn',  kindWarn && !kindError);
      tr.classList.toggle('scan-row-error', kindError);

      if (cells[1]) {
        cells[1].className = kindWarn ? 'file-link file-link-warn' : kindError ? 'file-link file-link-error' : 'file-link';
        cells[1].textContent = item.file_name || '';
        cells[1].title = kindWarn
          ? `경고 있음 · 클릭: 뷰어로 보기 · 더블클릭: 파일 열기`
          : '클릭: 뷰어로 보기 · 더블클릭: 파일 열기';
      }
      tr.style.cursor = 'pointer';
      tr.onclick = (e) => {
        // 체크박스·스핀버튼 클릭은 제외
        if (e.target.matches('input, button')) return;
        document.querySelectorAll('#scan-tbody tr.viewer-active')
          .forEach(r => r.classList.remove('viewer-active'));
        tr.classList.add('viewer-active');
        _requestPreview(item.kind);
      };
      tr.ondblclick = (e) => {
        if (e.target.matches('input, button')) return;
        if (item.file_path) bridge.openFile(item.file_path);
      };
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
    delete _previewData[kind];
    if (_currentKind === kind) _requestPreview(kind);
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
    const stateValue = (state.school_kind_override || '').trim();
    if (stateValue) return stateValue;

    const row = _el('school-kind-row');
    if (!row || row.style.display === 'none') return null;

    const selectValue = (_el('school-kind-select')?.value || '').trim();
    return selectValue || null;
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

    const spinEl = _el(`spin-${kind}`);
    const spinVal = parseInt(spinEl?.textContent || '', 10);
    const params = {
      kind,
      file_path:      item.file_path,
      sheet_name:     item.sheet_name     || '',
      header_row:     item.has_structured_rows === false ? null : (item.header_row ?? null),
      data_start_row: item.has_structured_rows === false ? null : (Number.isFinite(spinVal) && spinVal > 0 ? spinVal : (item.data_start_row ?? null)),
      issue_rows:     item.issue_rows     || [],
      start_row:      1,
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
      issue_rows, row_marks,
    } = payload;

    _previewData[kind] = {
      columns, rows, total_count, truncated,
      source_file, sheet_name,
      start_row:      payload.start_row ?? 1,
      header_row:     header_row     ?? null,
      data_start_row: data_start_row ?? null,
      issue_rows:     issue_rows     || [],
      row_marks:      row_marks      || null,
      actual_count:   payload.actual_count,
      displayed_count:payload.displayed_count,
      has_structured_rows: payload.has_structured_rows !== false,
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

    const structured = data.has_structured_rows !== false;
    const headerInfo = `헤더행: ${structured && data.header_row != null ? data.header_row : '-'}`;
    const startInfo  = ` | 시작행: ${structured && data.data_start_row != null ? data.data_start_row : '-'}`;
    const actualCount = Number.isFinite(data.actual_count) ? data.actual_count : (data.rows || []).length;
    const displayedCount = Number.isFinite(data.displayed_count) ? data.displayed_count : (data.rows || []).length;
    _el('preview-file-info').textContent =
      headerInfo + startInfo +
      ` | 실제 ${actualCount}행` +
      (data.truncated ? ` · ${displayedCount}행만 표시` : '');

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

  const columns = data.columns || [];
  const rawRows = data.rows || [];
  const previewStartRow = data.start_row ?? 1;
  const structured = data.has_structured_rows !== false;
  const dataStartRow = structured ? (data.data_start_row ?? 1) : null;
  const issueSet = structured ? new Set((data.row_marks?.warn_rows || data.row_marks?.issue_rows || [])) : new Set(); // row_marks 기반 — issue_rows fallback

  let lastVisibleIdx = -1;
  for (let i = rawRows.length - 1; i >= 0; i--) {
    const vals = (rawRows[i] || []).map(v => String(v || '').trim());
    if (vals.some(v => v)) {
      lastVisibleIdx = i;
      break;
    }
  }
  const rows = lastVisibleIdx >= 0 ? rawRows.slice(0, lastVisibleIdx + 1) : [];

  const nameCol  = columns.findIndex(h => ['성명','이름','학생이름'].some(k => h.includes(k)));
  const gradeCol = columns.findIndex(h => h.includes('학년'));
  const classCol = columns.findIndex(h => ['반','학급'].some(k => h.includes(k)) && !h.includes('학년'));

  const dupSet = new Set();
  if (nameCol >= 0) {
    const cnt = {};
    rows.forEach((r, i) => {
      const excelRow = previewStartRow + i;
      if (structured && excelRow < dataStartRow) return;
      const nm    = (r[nameCol] || '').replace(/[A-Z]+$/, '').trim();
      const grade = gradeCol >= 0 ? (r[gradeCol] || '') : '';
      const key   = `${grade}||${nm}`;
      if (nm) cnt[key] = (cnt[key] || []).concat(i);
    });
    Object.values(cnt).forEach(idxs => { if (idxs.length >= 2) idxs.forEach(i => dupSet.add(i)); });
  }

  const rowIssueMap = new Map();
  const addIssue = (excelRow, msg) => {
    if (!msg) return;
    const arr = rowIssueMap.get(excelRow) || [];
    if (!arr.includes(msg)) arr.push(msg);
    rowIssueMap.set(excelRow, arr);
  };

  rows.forEach((r, i) => {
    const excelRow = previewStartRow + i;
    if (structured && excelRow < dataStartRow) return;
    if (issueSet.has(excelRow)) addIssue(excelRow, '값 확인');
  });

  const mutedSet = new Set(data.row_marks?.muted_rows || []);
  const filtered = rows.reduce((acc, row, i) => {
    const excelRow = previewStartRow + i;
    const rowText = row.join(' ').toLowerCase();
    if (keyword && !rowText.includes(keyword)) return acc;
    const issues = rowIssueMap.get(excelRow) || [];
    const isPrestart = structured ? (mutedSet.size ? mutedSet.has(excelRow) : (excelRow < dataStartRow)) : false;
    const isIssue = !isPrestart && issues.length > 0;
    const isDup = !isPrestart && dupSet.has(i);
    if (issueOnly && !isIssue) return acc;
    if (dupOnly && !isDup) return acc;
    acc.push({ row, i, excelRow, issues, isPrestart, isIssue, isDup });
    return acc;
  }, []);

  const table = _el('preview-table');
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');

  const headers = [...columns, '비고'];
  const rowNoHeader = structured ? '#' : '';
  thead.innerHTML = `<tr><th style="width:40px;color:#94A3B8;font-weight:600;text-align:center">${rowNoHeader}</th>` +
    headers.map(h => `<th>${_esc(h)}</th>`).join('') + '</tr>';
  tbody.innerHTML = filtered.map(({ row, excelRow, issues, isPrestart, isIssue, isDup }) => {
    const finalCls = isPrestart ? 'row-prestart' : isIssue ? 'row-warn' : isDup ? 'row-dup' : '';
    const remark = isPrestart ? '' : (issues.length ? _esc(issues.join(', ')) : '');
    const cells = row.map(v => `<td>${_esc(v)}</td>`).join('') + `<td>${remark}</td>`;
    return `<tr class="${finalCls}"><td style="color:#94A3B8;font-size:11px;text-align:center;user-select:none">${structured ? excelRow : ''}</td>${cells}</tr>`;
  }).join('');

  const firstIssueExcelRow = [...rowIssueMap.keys()].sort((a, b) => a - b)[0] ?? null;
  return {
    issueCount: rowIssueMap.size,
    rowIssueMap,
    firstIssueExcelRow,
  };
}

function _syncPreviewWarn(kind, meta = {}) {
  const previewWarn = _el('preview-warn');
  if (!previewWarn) return;

  const data = _previewData[kind] || {};
  if (data.has_structured_rows === false) {
    previewWarn.innerHTML = `<div class="pw-inline">헤더를 찾지 못해 행번호를 표시하지 않습니다.</div>`;
    previewWarn.style.color = '#92400E';
    _refreshWarnUI();
    return;
  }
  const rowIssueMap = meta.rowIssueMap || new Map();
  const count = typeof meta.issueCount === 'number' ? meta.issueCount : rowIssueMap.size;

  if (!count) {
    previewWarn.textContent = '';
    previewWarn.style.color = '';
    _refreshWarnUI();
    return;
  }

  const firstEntry = rowIssueMap.entries().next().value;
  const excelRow = firstEntry?.[0];
  const firstReasons = firstEntry?.[1] || [];
  const lead = `${excelRow}행 ${firstReasons[0] || '확인'}`;
  const suffix = count > 1 ? ` 외 ${count - 1}건` : '';

  previewWarn.innerHTML = `<div class="pw-inline">⚠ 경고 · ${_esc(lead)}${_esc(suffix)}</div>`;
  previewWarn.style.color = '#92400E';

  _refreshWarnUI();
}

function _refreshWarnUI() {
  // events + status 둘 다 본다 — hold 포함
  const status = _lastScanData?.status || null;
  const events = _lastScanData?.events || [];
  StatusUI.renderBadge('scan-status-badge', status?.badge, '완료');
  const errMsgs  = events.filter(e => e.level === 'error').map(e => e.message);
  const holdMsgs = events.filter(e => e.level === 'hold').map(e => e.message);
  const warnMsgs = events.filter(e => e.level === 'warn').map(e => e.message);
  const combinedMsgs = [...warnMsgs, ...holdMsgs];
  if (errMsgs.length)           _showScanWarnCard(errMsgs,        'error', status);
  else if (combinedMsgs.length) _showScanWarnCard(combinedMsgs,   'warn',  null);
  else                          _hideScanWarnCard();
  _applyRowMarkStyles(_lastScanData?.row_marks || []);
}

  function _applyRowMarkStyles(rowMarks) {
    const kindError = new Set();
    const kindWarn  = new Set();
    const FILE_KEY_TO_KIND = {
      freshmen: '신입생', transfer_in: '전입생',
      transfer_out: '전출생', teachers: '교직원',
    };
    // item.severity 기반 — _applyScanTable이 이미 처리했으므로 현재 tr 클래스 기준으로 수집
    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      if (tr.classList.contains('scan-row-error')) kindError.add(kind);
      else if (tr.classList.contains('scan-row-warn')) kindWarn.add(kind);
    });
    // row_marks + events로 추가 보완 (기존에 없는 kind만)
    (rowMarks || []).forEach(m => {
      const kind = FILE_KEY_TO_KIND[m.file_key];
      if (!kind) return;
      if (m.level === 'error' && !kindError.has(kind)) kindError.add(kind);
      else if (m.level === 'warn' && !kindError.has(kind) && !kindWarn.has(kind)) kindWarn.add(kind);
    });
    (_lastScanData?.events || []).forEach(e => {
      const kind = FILE_KEY_TO_KIND[e.file_key];
      if (!kind) return;
      if (e.level === 'error' && !kindError.has(kind)) kindError.add(kind);
      else if (e.level === 'warn' && !kindError.has(kind) && !kindWarn.has(kind)) kindWarn.add(kind);
    });

    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      const isError = kindError.has(kind);
      const isWarn  = kindWarn.has(kind) && !isError;
      tr.classList.toggle('scan-row-warn',  isWarn);
      tr.classList.toggle('scan-row-error', isError);
      const fileCell = tr.querySelectorAll('td')[1];
      if (fileCell && fileCell.textContent) {
        fileCell.classList.toggle('file-link-warn',  isWarn);
        fileCell.classList.toggle('file-link-error', isError);
      }
    });
  }


function _normalizeScanCardMessage(msg, mode) {
  let s = String(msg || '').trim();
  if (!s) return '';
  s = s.replace(/^\[(WARN|ERROR)\]\s*/, '');
  if (/헤더를 찾을 수 없습니다|구조를 읽을 수 없습니다/.test(s)) {
    return '헤더를 찾을 수 없습니다.';
  }
  if (/시트가\s*\d+개 있습니다/.test(s) || /시트가 여러 개 있습니다/.test(s)) {
    return '시트가 여러 개 있습니다. 첫 번째 시트만 사용합니다.';
  }
  s = s.replace(/(\d+)행\s*'class'\s*값이\s*비어\s*있습니다\.?/gi, '$1행 반 정보가 비어 있습니다.');
  s = s.replace(/(\d+)행\s*'grade'\s*값이\s*비어\s*있습니다\.?/gi, '$1행 학년 정보가 비어 있습니다.');
  s = s.replace(/(\d+)행\s*'name'\s*값이\s*비어\s*있습니다\.?/gi, '$1행 이름 정보가 비어 있습니다.');
  s = s.replace(/'class'\s*값/gi, '반 정보');
  s = s.replace(/'grade'\s*값/gi, '학년 정보');
  s = s.replace(/'name'\s*값/gi, '이름 정보');
  return s;
}

function _showScanWarnCard(messages, mode = 'warn', status = null) {
  const el = _el('scan-warn-card');
  if (!el) return;
  UICommon.renderStatusCard(el, messages, mode, status);
}

  function _hideScanWarnCard() {
    const el = _el('scan-warn-card');
    if (!el) return;
    UICommon.hideStatusCard(el);
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

  function _showSchoolKindModal() {
    // 기존 인라인 UI 숨기기
    _setSchoolKindWarn(false);
    _setBadge('warn', '확인 필요');
    _setMessage('');
    App.setStepState(2, 'warn');

    const bd = document.createElement('div');
    bd.className = 'confirm-modal-backdrop';
    bd.innerHTML = `
      <div class="confirm-modal">
        <div class="confirm-modal-title">학교 구분 선택 필요</div>
        <div class="confirm-modal-body">
          학교명에서 학교 구분(초/중/고)을 자동으로 판별하지 못했습니다.<br>
          학교 구분을 선택한 뒤 다시 스캔해 주세요.
        </div>
        <div class="confirm-modal-options">
          <label class="confirm-modal-option selected">
            <input type="radio" name="school-kind-modal" value="초등부" checked>
            <div><div class="confirm-modal-option-label">초등부</div></div>
          </label>
          <label class="confirm-modal-option">
            <input type="radio" name="school-kind-modal" value="중등부">
            <div><div class="confirm-modal-option-label">중등부</div></div>
          </label>
          <label class="confirm-modal-option">
            <input type="radio" name="school-kind-modal" value="고등부">
            <div><div class="confirm-modal-option-label">고등부</div></div>
          </label>
          <label class="confirm-modal-option">
            <input type="radio" name="school-kind-modal" value="">
            <div><div class="confirm-modal-option-label">기타(빈칸)</div></div>
          </label>
        </div>
        <div class="confirm-modal-footer">
          <button class="btn-primary" id="school-kind-modal-ok" >선택 후 다시 스캔</button>
        </div>
      </div>`;
    document.body.appendChild(bd);

    bd.querySelectorAll('input[type="radio"]').forEach(r => {
      r.addEventListener('change', () => {
        bd.querySelectorAll('.confirm-modal-option').forEach(o => o.classList.remove('selected'));
        r.closest('.confirm-modal-option').classList.add('selected');
      });
    });

    bd.querySelector('#school-kind-modal-ok').addEventListener('click', () => {
      const selected = bd.querySelector('input[name="school-kind-modal"]:checked')?.value ?? '초등부';
      bd.remove();
      // state에 override 저장 + select UI 동기화
      state.school_kind_override = selected || null;
      const sel = _el('school-kind-select');
      if (sel) { sel.value = selected; }
      _setSchoolKindWarn(!!selected);
      // 직접 스캔 재실행 (btn.click()은 이벤트 루프 문제 있음)
      start();
    });
  }

  // ──────────────────────────────────────────────
  // 학년도 아이디 규칙 갱신 (StatusPanel)
  // ──────────────────────────────────────────────
  function _updateGradeMap(data) {
    const maxGrade = Number(data.grade_rule_max_grade);
    if (typeof Panel !== 'undefined' && Panel.setGradeCount) {
      if (Number.isFinite(maxGrade) && maxGrade > 0) Panel.setGradeCount(maxGrade);
      else Panel.setGradeCount(state.selected_school || data.school_name || '');
    }

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
          <button class="btn-primary" id="cm-date-confirm" >확인</button>
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
    _setBadge('running', '스캔 중');

    // 이전 스캔 결과 초기화
    _previewData = {};
    _currentKind = null;
    _lastScanData = null;
    _el('btn-run').disabled = true;
    App.setFloatingNext(false, null);
    if (typeof Panel !== 'undefined' && Panel.updateRosterMapBtn) Panel.updateRosterMapBtn(null);
    _hideSchoolKindWarn();
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
      _setBadge('err', '실패');
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

    const statusMsgs = UICommon.collectMessages({ status: _lastScanData?.status, events: _lastScanData?.events || [] });
    // FRESHMEN_NO_ROSTER_MANUAL 케이스는 수동 입력으로 대체 가능 — 차단하지 않음
    const isFreshmenManualMode = (_lastScanData?.events || []).some(e => e.code === 'FRESHMEN_NO_ROSTER_MANUAL');
    const hasStrictRosterError = !isFreshmenManualMode && statusMsgs.some(msg => String(msg || '').includes('학생명부가 필요합니다.'));
    if (hasStrictRosterError) return false;

    _lastScanData.grade_year_map = {
      ...(_lastScanData.grade_year_map || {}),
      ...(overrides || {}),
    };
    _lastScanData.can_execute = true;
    _lastScanData.can_execute_after_input = false;

    const removedCode = 'FRESHMEN_NO_ROSTER_MANUAL';
    const removedMessage = '학교 폴더에 명부를 추가하거나, 사이드바에서 학년도 아이디 규칙을 직접 입력하세요.';

    const nextEvents = Array.isArray(_lastScanData.events)
      ? _lastScanData.events.filter(e => e?.code !== removedCode && String(e?.message || '').trim() !== removedMessage)
      : [];
    _lastScanData.events = nextEvents;

    const prevStatus = _lastScanData?.status || {};
    const nextStatusMessages = Array.isArray(prevStatus.messages)
      ? prevStatus.messages.filter(m => m?.code !== removedCode && String(m?.text || '').trim() !== removedMessage)
      : [];

    const hasError = nextStatusMessages.some(m => m?.level === 'error') || nextEvents.some(e => e?.level === 'error');
    const hasHold = nextStatusMessages.some(m => m?.level === 'hold') || nextEvents.some(e => e?.level === 'hold');
    const hasWarn = nextStatusMessages.some(m => m?.level === 'warn') || nextEvents.some(e => e?.level === 'warn');
    const nextLevel = hasError ? 'error' : hasHold ? 'hold' : hasWarn ? 'warn' : 'ok';
    const nextBadge = nextLevel === 'error'
      ? { type: 'err', text: '오류' }
      : nextLevel === 'hold'
        ? { type: 'warn', text: '보류' }
        : nextLevel === 'warn'
          ? { type: 'warn', text: '경고' }
          : { type: 'ok', text: '완료' };
    const nextSummary = nextLevel === 'error'
      ? `오류 ${nextStatusMessages.length}건이 있습니다.`
      : nextLevel === 'hold'
        ? `보류 ${nextStatusMessages.length}건이 있습니다.`
        : nextLevel === 'warn'
          ? `경고 ${nextStatusMessages.length}건이 있습니다.`
          : '완료';

    _lastScanData.status = {
      ...prevStatus,
      level: nextLevel,
      badge: nextBadge,
      messages: nextStatusMessages,
      detail_messages: nextStatusMessages.map(m => String(m?.text || '').trim()).filter(Boolean),
      summary_text: nextSummary,
      action_text: nextLevel === 'ok' ? '' : (prevStatus.action_text || ''),
    };

    if (nextLevel === 'error') _showScanWarnCard(nextStatusMessages.map(m => m?.text || ''), 'error', _lastScanData.status);
    else if (nextLevel === 'hold' || nextLevel === 'warn') _showScanWarnCard(nextStatusMessages.map(m => m?.text || ''), 'warn', _lastScanData.status);
    else _hideScanWarnCard();

    _setBadge(nextBadge.type, nextBadge.text);
    _setMessage('학년도 아이디 규칙이 적용되었습니다. 바로 실행할 수 있습니다.');
    _el('btn-run').disabled = false;
    App.setFloatingNext(true, 'run');

    const previewWarn = _el('preview-warn');
    if (previewWarn && previewWarn.textContent.includes('학년도 아이디 규칙')) {
      previewWarn.textContent = '';
    }

    _updateGradeMap(_lastScanData);
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

    _setBadge('idle', '대기');
    _setMessage('');
    _hideSchoolKindWarn();
    _hideScanWarnCard();

    // 스캔 표 초기화
    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      tr.classList.remove('scan-row-warn', 'scan-row-error', 'viewer-active');
      const cells = tr.querySelectorAll('td');
      [1, 2].forEach(c => { if (cells[c]) cells[c].textContent = ''; });
      if (cells[1]) {
        cells[1].classList.remove('file-link', 'file-link-warn', 'file-link-error');
        cells[1].onclick = null;
        cells[1].ondblclick = null;
        cells[1].title = '';
      }
      tr.onclick = null;
      tr.ondblclick = null;
      tr.style.cursor = '';
      const spinVal = _el(`spin-${kind}`);
      if (spinVal) { spinVal.textContent = '-'; spinVal.style.color = '#94A3B8'; }
      const chk = _el(`chk-${kind}`);
      if (chk) { chk.checked = false; chk.disabled = false; }
    });

    // 뷰어
    if (_viewerOpen) toggleViewer();
    _el('preview-file-info').textContent = '헤더행: - | 시작행: - | 실제 -행';
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
