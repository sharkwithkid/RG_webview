/**
 * diff_tab.js — 명단 비교 탭 로직
 */
'use strict';
const Diff = (() => {
  const COMPARE_KIND = '재학생';
  let _lastScanData = null;
  let _compareItem = null;
  let _previewData = null;
  let _viewerOpen = false;
  let _bound = false;
  function _bindOnce() {
    if (_bound) return;
    _bound = true;
    document.addEventListener('click', (e) => {
      const minus = e.target.closest('#diff-spin-minus');
      const plus = e.target.closest('#diff-spin-plus');
      const file = e.target.closest('#diff-compare-file-name');
      if (minus) {
        e.preventDefault();
        _spin(-1);
      }
      if (plus) {
        e.preventDefault();
        _spin(1);
      }
      if (file && _compareItem?.file_path) {
        _requestPreview();
      }
    });
    document.addEventListener('dblclick', (e) => {
      const file = e.target.closest('#diff-compare-file-name');
      if (file && _compareItem?.file_path) bridge.openFile(_compareItem.file_path);
    });
  }
  async function scan() {
    _bindOnce();
    if (state.isDiffScanning || state.isDiffRunning) return;
    if (!state.work_root)       { toast('작업 폴더가 설정되지 않았습니다.', 'warn'); return; }
    if (!state.selected_school) { toast('학교를 먼저 선택해 주세요.', 'warn'); return; }
    state.isDiffScanning = true;
    _el('btn-scan-diff').disabled = true;
    _el('btn-run-diff').disabled = true;
    _setScanBadge('running', '스캔 중');
    _setScanMessage('재학생 파일 구조를 확인하고 있습니다.');
    _setRunInfo('스캔 완료 후 명단 비교를 실행할 수 있습니다.');
    _setSummary('재학생 파일의 시작행을 확인해 주세요.');
    _clearTables();
    _clearFiles();
    _setStats('-', '-', '-', '-', '-', '-');
    _resetCompareRow();
    const params = {
      work_root:         state.work_root,
      school_name:       state.selected_school,
      school_start_date: state.school_start_date,
      work_date:         state.work_date,
      roster_xlsx:       state.roster_log_path || '',
      col_map:           state.roster_col_map  || {},
    };
    const res = JSON.parse(await bridge.startScanDiff(JSON.stringify(params)));
    if (!res.ok) {
      state.isDiffScanning = false;
      _el('btn-scan-diff').disabled = false;
      _setScanBadge('err', '오류');
      _setScanMessage(res.error || '명단 비교 스캔 시작 실패');
    }
  }
  async function run() {
    _bindOnce();
    if (state.isDiffScanning || state.isDiffRunning) return;
    if (!_lastScanData || !_compareItem) {
      toast('먼저 파일 내용 스캔을 완료해 주세요.', 'warn');
      return;
    }
    const chk = _el('chk-재학생');
    if (!chk?.checked) {
      toast('재학생 파일의 시작행을 확인해 주세요.', 'warn', 3500);
      return;
    }
    const basisDate = _lastScanData.roster_basis_date || '';
    if (_lastScanData.roster_date_mismatch && basisDate) {
      _handleDiffRosterDateMismatch(basisDate);
      return;
    }
    await _runWithBasisDate(basisDate);
  }
  async function _runWithBasisDate(basisDate) {
    state.isDiffRunning = true;
    _el('btn-run-diff').disabled = true;
    _setRunBadge('running', '실행 중');
    _setRunInfo('명단 비교를 진행하고 있습니다.');
    const params = {
      work_root:         state.work_root,
      school_name:       state.selected_school,
      school_start_date: state.school_start_date,
      work_date:         state.work_date,
      roster_basis_date: basisDate || '',
      roster_xlsx:       state.roster_log_path || '',
      col_map:           state.roster_col_map  || {},
      layout_overrides:  _getLayoutOverrides(),
    };
    const res = JSON.parse(await bridge.startRunDiff(JSON.stringify(params)));
    if (!res.ok) {
      state.isDiffRunning = false;
      _el('btn-run-diff').disabled = false;
      _setRunBadge('err', '오류');
      _setRunInfo(res.error || '명단 비교 시작 실패');
    }
  }
  function onScanFinished(data) {
    state.isDiffScanning = false;
    state.last_diff_logs = data.logs || [];
    _lastScanData = data || null;
    _compareItem = (data.items || []).find(i => i.kind === COMPARE_KIND) || null;
    if (!data.ok || !_compareItem) {
      const status = data.status || null;
      const errMsg = (status?.messages || []).find(m => m.level === 'error')?.text
        || '재학생 파일을 확인할 수 없습니다.';
      _setScanBadge('err', '오류');
      _setScanMessage('');
      _showWarnCard('diff-scan-warn-card', [errMsg], 'error', status);
      _applyCompareWarnStyles('error');
      _el('btn-scan-diff').disabled = false;
      return;
    }
    _applyCompareRow(_compareItem);

    // 뱃지 + 카드 — status 하나만 본다
    const status = data.status || null;
    StatusUI.renderBadge('diff-scan-status-badge', status?.badge, '완료');
    _setScanMessage('');
    const errMsgs  = (status?.messages || []).filter(m => m.level === 'error').map(m => m.text);
    const warnMsgs = (status?.messages || []).filter(m => m.level === 'warn').map(m => m.text);
    const holdMsgs = (status?.messages || []).filter(m => m.level === 'hold').map(m => m.text);
    if (errMsgs.length) {
      _showWarnCard('diff-scan-warn-card', errMsgs, 'error', status);
      _applyCompareWarnStyles('error');
    } else if (holdMsgs.length || warnMsgs.length) {
      _showWarnCard('diff-scan-warn-card', [...holdMsgs, ...warnMsgs], 'warn', status);
      _applyCompareWarnStyles('warn');
    } else {
      _hideWarnCard('diff-scan-warn-card');
      _applyCompareWarnStyles('ok');
    }
    _setRunInfo('스캔 완료. 시작행 확인 후 명단 비교를 실행해 주세요.');
    _el('btn-scan-diff').disabled = false;
    _el('btn-run-diff').disabled = false;
    _requestPreview();
  }
  function onScanFailed(error) {
    state.isDiffScanning = false;
    _el('btn-scan-diff').disabled = false;
    _setScanBadge('err', '오류');
    _setScanMessage(error || '미리보기를 불러올 수 없습니다.');
  }
  function _handleDiffRosterDateMismatch(basisDate) {
    const workDate = state.work_date;
    const backdrop = document.createElement('div');
    backdrop.className = 'confirm-modal-backdrop';
    backdrop.innerHTML = `
      <div class="confirm-modal">
        <div class="confirm-modal-title">명부 기준일 설정</div>
        <div class="confirm-modal-body">학생명부 마지막 수정일과 작업일이 다릅니다.<br>어느 날짜를 명부 기준일로 사용할까요?</div>
        <div class="confirm-modal-options">
          <label class="confirm-modal-option selected">
            <input type="radio" name="diff-cm-date" value="basis" checked>
            <div><div class="confirm-modal-option-label">수정일 사용 — ${basisDate}</div><div class="confirm-modal-option-desc">명부 파일의 마지막 수정일을 기준으로 합니다.</div></div>
          </label>
          <label class="confirm-modal-option">
            <input type="radio" name="diff-cm-date" value="work">
            <div><div class="confirm-modal-option-label">작업일 사용 — ${workDate}</div><div class="confirm-modal-option-desc">오늘 작업일을 기준으로 진행합니다.</div></div>
          </label>
        </div>
        <div class="confirm-modal-footer"><button class="btn-primary" id="diff-cm-date-confirm" style="height:36px;padding:0 20px">확인</button></div>
      </div>`;
    document.body.appendChild(backdrop);
    backdrop.querySelectorAll('input[type="radio"]').forEach(radio => {
      radio.addEventListener('change', () => {
        backdrop.querySelectorAll('.confirm-modal-option').forEach(opt => opt.classList.remove('selected'));
        radio.closest('.confirm-modal-option').classList.add('selected');
      });
    });
    backdrop.querySelector('#diff-cm-date-confirm').addEventListener('click', async () => {
      const selected = backdrop.querySelector('input[name="diff-cm-date"]:checked')?.value;
      backdrop.remove();
      await _runWithBasisDate(selected === 'work' ? workDate : basisDate);
    });
  }
  function onFinished(data) {
    state.isDiffRunning = false;
    state.isDiffScanning = false;
    _el('btn-run-diff').disabled = false;
    if (!data.ok) {
      const errMsg = (data.status?.messages || []).find(m => m.level === 'error')?.text || '실패';
      _setRunBadge('err', '오류');
      _setRunInfo('');
      _showWarnCard('diff-run-warn-card', [errMsg], 'error', data.status || null);
      return;
    }
    const rosterOnlyCount = Number(data.roster_only_count || 0);
    const matchedCount = Number(data.matched_count || 0);
    const compareOnlyCount = Number(data.compare_only_count || 0);
    const unresolvedCount = Number(data.unresolved_count || 0);
    if (data.status?.badge) StatusUI.renderBadge('diff-run-status-badge', data.status.badge, '완료');
    else _setRunBadge(unresolvedCount > 0 ? 'warn' : 'ok', unresolvedCount > 0 ? '경고' : '완료');
    _setRunInfo('명단 비교 결과를 아래에서 확인해 주세요.');
    
    // 카드 — status 하나만 본다 (logs 파싱 없음)
    const runMsgs = (data.status?.messages || []).filter(m => ['warn','hold'].includes(m.level)).map(m => m.text);
    if (runMsgs.length) _showWarnCard('diff-run-warn-card', runMsgs, 'warn', data?.status || null);
    else _hideWarnCard('diff-run-warn-card');

    _setStats(
      `${rosterOnlyCount}명`,
      `${matchedCount}명`,
      `${compareOnlyCount}명`,
      `${unresolvedCount}명`,
      `${compareOnlyCount}명`,
      `${rosterOnlyCount}명`,
    );
    _renderSimpleRows('diff-tbody-roster-only', data.roster_only_rows || [], 3);
    _renderSimpleRows('diff-tbody-compare-only', data.compare_only_rows || [], 3);
    _renderUnresolvedRows('diff-tbody-unresolved', data.unresolved_rows || []);
    _setSummary('명단 비교 결과를 아래에서 확인해 주세요.');
    const filesEl = _el('diff-output-files');
    if (filesEl) {
      const first = (data.output_files || [])[0];
      filesEl.innerHTML = first
        ? `<div class="diff-output-link"><span class="lbl">비교 결과 파일</span><span class="lnk" onclick="Run.openFile('${_escJs(first.path)}')">${_escHtml(first.name)}</span></div>`
        : '<span class="muted">생성된 파일이 없습니다.</span>';
    }
  }
  function onFailed(error) {
    state.isDiffRunning = false;
    state.isDiffScanning = false;
    _el('btn-run-diff').disabled = false;
    _setRunBadge('err', '오류');
    _setRunInfo('');
    _showWarnCard('diff-run-warn-card', [`예기치 못한 오류: ${error}`], 'error');
    // 토스트 제거 — 카드가 있어서 중복
  }
  function onPreviewLoaded(payload) {
    _previewData = payload;
    const actualCount = Number.isFinite(payload.actual_count) ? payload.actual_count : (payload.rows || []).length;
    const displayedCount = Number.isFinite(payload.displayed_count) ? payload.displayed_count : (payload.rows || []).length;
    const structured = payload.has_structured_rows !== false;
    const info = `파일: ${payload.source_file || '-'} | 시트: ${payload.sheet_name || '-'} | 헤더행: ${structured && payload.header_row != null ? payload.header_row : '-'} | 시작행: ${structured && payload.data_start_row != null ? payload.data_start_row : '-'} | 실제 ${actualCount}행`;
    _el('diff-preview-file-info').textContent = info;
    if (payload.has_structured_rows === false) _el('diff-preview-warn').textContent = '헤더를 찾지 못해 행번호를 표시하지 않습니다.';
    else _el('diff-preview-warn').textContent = payload.truncated ? `${displayedCount}행만 표시합니다.` : '';
    const table = _el('diff-preview-table');
    if (!table) return;
    const cols = payload.columns || [];
    const rows = payload.rows || [];
    const startRow = payload.start_row || 1;
    const structured2 = payload.has_structured_rows !== false;
    const issueSet = structured2 ? new Set(payload.row_marks?.issue_rows || payload.row_marks?.warn_rows || payload.issue_rows || []) : new Set();
    const mutedSet = new Set(payload.row_marks?.muted_rows || []);
    const dataStart = structured2 ? (payload.data_start_row || 1) : null;
    table.querySelector('thead').innerHTML = `<tr><th style="width:40px;color:#94A3B8;font-weight:600;text-align:center">${structured2 ? '#' : ''}</th>` + cols.map(c => `<th>${_escHtml(c)}</th>`).join('') + '</tr>';
    if (!rows.length) {
      table.querySelector('tbody').innerHTML = `<tr><td colspan="${Math.max(cols.length+1,1)}" class="diff-empty">표시할 데이터가 없습니다.</td></tr>`;
    } else {
      table.querySelector('tbody').innerHTML = rows.map((r, i) => {
        const excelRow = startRow + i;
        const cls = structured2 ? ((mutedSet.size ? mutedSet.has(excelRow) : (excelRow < dataStart)) ? 'row-prestart' : (issueSet.has(excelRow) ? 'row-warn' : '')) : '';
        return `<tr class="${cls}"><td style="color:#94A3B8;font-size:11px;text-align:center;user-select:none">${structured2 ? excelRow : ''}</td>${r.map(v => `<td>${_escHtml(v)}</td>`).join('')}</tr>`;
      }).join('');
    }
    if (_viewerOpen) _el('diff-viewer-body').style.display = 'block';
  }
  function onPreviewFailed(kind, error) {
    _el('diff-preview-warn').textContent = error || '미리보기를 불러올 수 없습니다.';
  }
  async function _requestPreview() {
    if (!_compareItem?.file_path) return;
    if (state.isPreviewLoading) return;
    state.isPreviewLoading = true;
    _el('diff-preview-warn').textContent = '미리보기를 불러오는 중입니다.';
    const params = {
      kind: 'compare',
      file_path: _compareItem.file_path,
      sheet_name: _compareItem.sheet_name || '',
      header_row: _compareItem.has_structured_rows === false ? null : (_compareItem.header_row ?? null),
      data_start_row: _getCompareStartRow(),
      issue_rows: _compareItem.issue_rows || [],
      start_row: 1,
    };
    const res = JSON.parse(await bridge.startPreview(JSON.stringify(params)));
    if (!res.ok) {
      state.isPreviewLoading = false;
      _el('diff-preview-warn').textContent = res.error || '미리보기 시작 실패';
    }
  }
  function _applyCompareRow(item) {
    const nameEl = _el('diff-compare-file-name');
    nameEl.textContent = item.file_name || '-';
    nameEl.style.color = 'var(--text)';
    _el('diff-compare-sheet').textContent = item.sheet_name || '';
    const spin = _el('diff-spin-compare');
    if (spin) spin.textContent = String(item.data_start_row || '-');
    const chk = _el('chk-재학생');
    if (chk) { chk.checked = false; chk.disabled = false; }
  }
  function _resetCompareRow() {
    _lastScanData = null;
    _compareItem = null;
    _previewData = null;
    const nameEl = _el('diff-compare-file-name');
    nameEl.textContent = '-';
    nameEl.style.color = '';
    _el('diff-compare-sheet').textContent = '-';
    _el('diff-spin-compare').textContent = '-';
    const chk = _el('chk-재학생');
    if (chk) { chk.checked = false; chk.disabled = true; }
    _el('diff-preview-file-info').textContent = '파일: - | 시트: - | 헤더행: - | 시작행: -';
    _el('diff-preview-warn').textContent = '파일을 선택하면 전체 미리보기가 표시됩니다.';
    const table = _el('diff-preview-table');
    if (table) {
      table.querySelector('thead').innerHTML = '';
      table.querySelector('tbody').innerHTML = '';
    }
  }
  function _spin(delta) {
    const el = _el('diff-spin-compare');
    if (!el) return;
    const cur = parseInt(el.textContent, 10);
    if (Number.isNaN(cur)) return;
    const next = cur + delta;
    if (next < 1) return;
    el.textContent = String(next);
    const chk = _el('chk-재학생');
    if (chk) chk.checked = false;
    _requestPreview();
  }
  function _getCompareStartRow() {
    const v = parseInt(_el('diff-spin-compare')?.textContent || '', 10);
    return Number.isFinite(v) && v > 0 ? v : (_compareItem?.data_start_row || 2);
  }
  function _getLayoutOverrides() {
    return { compare: { data_start_row: _getCompareStartRow() } };
  }
  function showLog() { showLogDialog('명단 비교 로그', state.last_diff_logs); }
  function toggleViewer() {
    _viewerOpen = !_viewerOpen;
    _el('diff-viewer-body').style.display = _viewerOpen ? 'block' : 'none';
    const btn = _el('btn-toggle-diff-viewer');
    if (btn) btn.textContent = _viewerOpen ? '접기 ▴' : '펼치기 ▾';
  }
  function reset() {
    _bindOnce();
    _viewerOpen = false;
    _el('diff-viewer-body').style.display = 'none';
    const btn = _el('btn-toggle-diff-viewer');
    if (btn) btn.textContent = '펼치기 ▾';
    _setScanBadge('idle', '대기');
    _setRunBadge('idle', '대기');
    _setScanMessage('재학생 파일 구조를 먼저 확인해 주세요.');
    _setRunInfo('스캔을 통과한 후 명단 비교를 실행하고 결과를 확인합니다.');
    _hideWarnCard('diff-scan-warn-card');
    _hideWarnCard('diff-run-warn-card');
    _setSummary('명단 비교 실행 버튼을 눌러 주세요.');
    _clearFiles();
    _clearTables();
    _setStats('-', '-', '-', '-', '-', '-');
    _el('btn-scan-diff').disabled = false;
    _el('btn-run-diff').disabled = true;
    _resetCompareRow();
  }
  function _setScanBadge(type, text) {
    const el = _el('diff-scan-status-badge');
    if (!el) return;
    el.className = `status-badge badge-${type}`;
    el.textContent = text;
  }
  function _setRunBadge(type, text) {
    const el = _el('diff-run-status-badge');
    if (!el) return;
    el.className = `status-badge badge-${type}`;
    el.textContent = text;
  }
  function _setScanMessage(text) { const el = _el('diff-scan-message'); if (el) el.textContent = text || ''; }
  function _setRunInfo(text) { const el = _el('diff-run-info'); if (el) el.textContent = text || ''; }
  function _setSummary(text) { const el = _el('diff-summary'); if (el) el.textContent = text || ''; }
  function _setStats(rosterOnly, matched, compareOnly, unresolved, transferIn, transferOut) {
    const map = {
      'diff-stat-roster-only': rosterOnly,
      'diff-stat-matched': matched,
      'diff-stat-compare-only': compareOnly,
      'diff-stat-unresolved': unresolved,
      'diff-stat-transfer-in': transferIn,
      'diff-stat-transfer-out': transferOut,
    };
    Object.entries(map).forEach(([id, value]) => { const el = _el(id); if (el) el.textContent = value; });
  }
  function _renderSimpleRows(tbodyId, rows, colCount) {
    const tbody = _el(tbodyId); if (!tbody) return;
    const viewRows = Array.isArray(rows) ? rows : [];
    if (!viewRows.length) { tbody.innerHTML = `<tr><td colspan="${colCount}" class="diff-empty">해당 항목 없음</td></tr>`; return; }
    tbody.innerHTML = viewRows.map(r => `<tr><td>${_escHtml(r.grade)}</td><td>${_escHtml(r.class)}</td><td>${_escHtml(r.name)}</td></tr>`).join('');
  }
  function _renderUnresolvedRows(tbodyId, rows) {
    const tbody = _el(tbodyId); if (!tbody) return;
    const viewRows = Array.isArray(rows) ? rows : [];
    if (!viewRows.length) { tbody.innerHTML = '<tr><td colspan="4" class="diff-empty">해당 항목 없음</td></tr>'; return; }
    tbody.innerHTML = viewRows.map(r => `<tr><td>${_escHtml(r.grade)}</td><td>${_escHtml(r.class)}</td><td>${_escHtml(r.name)}</td><td>${_escHtml(r.hold_reason || r.reason || '')}</td></tr>`).join('');
  }
  function _makeEmptyWarnState() {
    // 하위 호환용 — 실제 판정은 status 기반
    return { hasWarn: false, hasError: false, messages: [], errorMessages: [], issueRows: new Set() };
  }
  function _applyCompareWarnStyles(mode) {
    const row = _el('diff-compare-row');
    const file = _el('diff-compare-file-name');
    if (!row || !file) return;
    row.classList.remove('scan-row-warn', 'scan-row-error');
    file.classList.remove('file-link-warn', 'file-link-error');
    if (mode === 'warn') {
      row.classList.add('scan-row-warn');
      if (file.textContent && file.textContent !== '-') file.classList.add('file-link-warn');
    } else if (mode === 'error') {
      row.classList.add('scan-row-error');
      if (file.textContent && file.textContent !== '-') file.classList.add('file-link-error');
    }
  }
  function _showWarnCard(id, messages, mode = 'warn', status = null) {
    const el = _el(id);
    if (!el) return;
    const html = (typeof StatusUI !== 'undefined' && StatusUI.normalizeStatusCard)
      ? StatusUI.normalizeStatusCard(messages, mode, status)
      : null;
    if (!html) {
      _hideWarnCard(id);
      return;
    }
    el.classList.toggle('error', mode === 'error');
    el.innerHTML = html;
    el.style.display = 'block';
  }
  function _hideWarnCard(id) {
    const el = _el(id);
    if (!el) return;
    el.style.display = 'none';
    el.innerHTML = '';
    el.classList.remove('error');
  }
  function _clearTables() {
    _renderSimpleRows('diff-tbody-roster-only', [], 3);
    _renderSimpleRows('diff-tbody-compare-only', [], 3);
    _renderUnresolvedRows('diff-tbody-unresolved', []);
  }
  function _clearFiles() { const filesEl = _el('diff-output-files'); if (filesEl) filesEl.innerHTML = '<span class="muted">실행 전</span>'; }
  function _escHtml(v) { return String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
  function _escJs(v) { return String(v ?? '').replace(/\\/g,'\\\\').replace(/'/g, "\\'"); }
  return { scan, run, onScanFinished, onScanFailed, onFinished, onFailed, onPreviewLoaded, onPreviewFailed, showLog, reset, toggleViewer };
})();