/**
 * run_tab.js — 실행·결과 탭 + Diff 탭 로직
 *
 * 의존: main.js (state, bridge, App, _el, showLogDialog)
 *       scan_tab.js (Scan)
 *       status_panel.js (Panel)
 * HTML ID: btn-run, run-status-badge, run-info, run-hold-warn,
 *           result-note, btn-goto-notice,
 *           sum-school, sum-year, sum-freshmen, sum-teacher,
 *           sum-transfer, sum-withdraw, sum-transfer-check, sum-withdraw-check,
 *           output-file-list, sheet-tabs, run-table, run-search,
 *           btn-run-dup, run-preview-info,
 *           btn-run-diff, diff-status-badge, diff-target-year,
 *           diff-result-info, diff-output-files
 */

'use strict';

// ─────────────────────────────────────────────────
// Run 탭
// ─────────────────────────────────────────────────
const Run = (() => {

  let _outputFiles  = [];   // [{ name, path }]
  let _currentFile  = null;
  let _sheetData    = {};   // { sheetName: { headers, rows } }
  let _currentSheet = null;
  let _pendingSheets = [];  // 순차 로드 대기 시트 큐
  let _dupOnly      = false;
  let _lastRunData  = null;

  // ──────────────────────────────────────────────
  // 실행 시작
  // ──────────────────────────────────────────────
  async function start() {
    if (state.isRunning) return;

    const layoutOverrides    = Scan.getLayoutOverrides();
    const schoolKindOverride = Scan.getSchoolKindOverride();

    const params = {
      work_date:            state.work_date,
      school_start_date:    state.school_start_date,
      layout_overrides:     Object.keys(layoutOverrides).length ? layoutOverrides : null,
      school_kind_override: schoolKindOverride,
    };

    state.isRunning = true;
    _el('btn-run').disabled = true;
    _setBadge('running', '실행 중');
    _el('run-info').textContent = '실행 중...';
    _el('run-hold-warn').style.display = 'none';

    const res = JSON.parse(await bridge.startRunMain(JSON.stringify(params)));
    if (!res.ok) {
      state.isRunning = false;
      _el('btn-run').disabled = false;
      _setBadge('err', '실행 실패');
      _el('run-info').textContent = res.error || '실행 시작 실패';
    }
    // 비동기 완료 → main.js → bridge.runFinished → onFinished / onFailed
  }

  // ──────────────────────────────────────────────
  // 실행 완료 (main.js에서 호출)
  // ──────────────────────────────────────────────
  function onFinished(data) {
    _el('btn-run').disabled = false;
    _lastRunData = data;

    if (!data.ok) {
      const err = (data.logs || []).find(l => l.level === 'error');
      _setBadge('err', '실행 실패');
      _el('run-info').textContent = err ? err.message : '실행 중 오류가 발생했습니다.';
      App.setStepState(3, 'warn');
      toast('작업 실행 중 오류 — 실행 로그 보기에서 확인하세요.', 'err');
      return;
    }

    // 스텝 상태
    App.setStepState(3, 'done');
    App.setStepState(4, 'active');

    // 경고 / 완료 뱃지
    const warnLogs = (data.logs || []).filter(l => l.level === 'warn');
    if (warnLogs.length) {
      _setBadge('warn', '경고');
      _el('run-info').textContent = `실행 완료 — 경고 ${warnLogs.length}건: ${warnLogs[0].message}`;
    } else {
      _setBadge('ok', '실행 완료');
      _el('run-info').textContent = '실행 완료';
    }

    // 보류 경고
    const realHold = (data.transfer_in_hold || 0) + (data.transfer_out_hold || 0)
                   - (data.transfer_out_auto_skip || 0);
    const autoSkip = data.transfer_out_auto_skip || 0;
    const holdWarn = _el('run-hold-warn');
    if (realHold > 0) {
      let msg = `보류 ${realHold}건이 있습니다. 생성된 파일의 보류 시트를 확인해 주세요.`;
      if (autoSkip > 0) msg += ` (자동제외 ${autoSkip}건 별도)`;
      holdWarn.textContent   = msg;
      holdWarn.style.display = '';
    } else {
      holdWarn.style.display = 'none';
    }

    // 상태 요약 그리드
    _updateSummary(data);

    // 출력 파일 목록
    _outputFiles = data.output_files || [];
    _renderOutputFiles();

    // 명단 기록 버튼 활성
    state.pending_roster_log = !!state.roster_log_path;
    Panel.setRosterBtns(!!state.roster_log_path, !!state.roster_log_path);

    // 안내문 탭으로 이동 버튼 표시
    _el('btn-goto-notice').style.display = '';
  }

  function onFailed(error) {
    _el('btn-run').disabled = false;
    _setBadge('err', '실행 실패');
    _el('run-info').textContent = '예기치 못한 오류가 발생했습니다.';
    App.setStepState(3, 'warn');
    toast('실행 오류 — 실행 로그 보기에서 자세한 내용을 확인하세요.', 'err');
  }

  // ──────────────────────────────────────────────
  // 상태 요약 그리드
  // ──────────────────────────────────────────────
  function _updateSummary(data) {
    const s        = v => (v != null && v !== '') ? String(v) : '-';
    const autoSkip = data.transfer_out_auto_skip || 0;
    _el('sum-school').textContent         = s(state.selected_school);
    _el('sum-year').textContent           = s(state.school_start_date?.slice(0, 4));
    _el('sum-freshmen').textContent       = s(data.freshmen_count);
    _el('sum-teacher').textContent        = s(data.teacher_count);
    _el('sum-transfer').textContent       = s(data.transfer_in_done);
    _el('sum-withdraw').textContent       = s(
      autoSkip > 0
        ? `${data.transfer_out_done} (자동제외 ${autoSkip}건)`
        : data.transfer_out_done
    );
    _el('sum-transfer-check').textContent = s(data.transfer_in_done);
    _el('sum-withdraw-check').textContent = s(data.transfer_out_done);
  }

  // ──────────────────────────────────────────────
  // 출력 파일 목록
  // ──────────────────────────────────────────────
  function _renderOutputFiles() {
    const el = _el('output-file-list');
    if (!el) return;
    el.textContent = '';

    if (!_outputFiles.length) {
      const empty = document.createElement('span');
      empty.className   = 'muted';
      empty.textContent = '생성된 파일 없음';
      el.appendChild(empty);
      return;
    }

    _outputFiles.forEach(f => {
      const row  = document.createElement('div');
      row.className = 'output-file-item';
      const link = document.createElement('span');
      link.className   = 'output-file-name';
      link.textContent = f.name;
      link.addEventListener('click', () => Run.openFile(f.path));
      row.appendChild(link);
      el.appendChild(row);
    });

    // 첫 파일 자동 뷰어 로드
    _loadFileViewer(_outputFiles[0].path);
  }

  // ──────────────────────────────────────────────
  // 파일 뷰어 (openpyxl → bridge.startPreview 대신
  //   run 결과 파일은 직접 openFile로 열거나
  //   별도 preview 슬롯이 필요 — 현재는 시트탭 구조만 유지)
  // ──────────────────────────────────────────────
  async function _loadFileViewer(filePath) {
    if (!filePath) return;
    _currentFile  = filePath;
    _sheetData    = {};
    _currentSheet = null;

    // 시트탭 초기화
    const tabs = _el('sheet-tabs');
    if (tabs) tabs.innerHTML = '';
    const table = _el('run-table');
    if (table) {
      table.querySelector('thead').innerHTML = '';
      table.querySelector('tbody').innerHTML = '';
    }
    _el('run-preview-info').textContent = '불러오는 중...';

    // xlsx 시트 목록 먼저 파악
    let sheetNames = [];
    try {
      const metaRes = JSON.parse(
        await bridge.readXlsxMeta(filePath, '', 1)
      );
      if (metaRes.ok) sheetNames = metaRes.data.sheets || [];
    } catch (e) { /* fallback: 첫 시트만 */ }

    if (!sheetNames.length) sheetNames = [''];

    // 첫 시트 로드 요청 (완료 후 onPreviewLoaded에서 다음 시트 순차 로드)
    _pendingSheets = sheetNames.slice(1);  // 나머지 시트 큐
    state.isPreviewLoading = true;
    const params = {
      kind:           'run_output',
      file_path:      filePath,
      sheet_name:     sheetNames[0],
      header_row:     1,
      data_start_row: 2,
    };
    const res = JSON.parse(await bridge.startPreview(JSON.stringify(params)));
    if (!res.ok) {
      state.isPreviewLoading = false;
      _el('run-preview-info').textContent = res.error || '미리보기 시작 실패';
    }
  }

  function onPreviewLoaded(payload) {
    const { sheet_name, columns, rows } = payload;
    _sheetData[sheet_name] = { headers: columns, rows };

    // 시트탭 추가
    const tabs = _el('sheet-tabs');
    if (tabs && !tabs.querySelector(`[data-sheet="${sheet_name}"]`)) {
      const btn = document.createElement('button');
      btn.className       = 'sheet-tab';
      btn.dataset.sheet   = sheet_name;
      btn.textContent     = sheet_name;
      btn.onclick         = () => _switchSheet(sheet_name);
      tabs.appendChild(btn);
    }

    if (!_currentSheet) _switchSheet(sheet_name);

    // 다음 시트 순차 로드
    if (_pendingSheets.length) {
      const next = _pendingSheets.shift();
      state.isPreviewLoading = true;
      bridge.startPreview(JSON.stringify({
        kind:           'run_output',
        file_path:      _currentFile,
        sheet_name:     next,
        header_row:     1,
        data_start_row: 2,
      })).then(res => {
        const r = JSON.parse(res);
        if (!r.ok) state.isPreviewLoading = false;
      });
    }
  }

  function _switchSheet(name) {
    _currentSheet = name;
    document.querySelectorAll('.sheet-tab').forEach(b =>
      b.classList.toggle('active', b.dataset.sheet === name)
    );
    _renderRunTable();
  }

  function _renderRunTable() {
    const data = _sheetData[_currentSheet];
    if (!data) return;

    const keyword = (_el('run-search')?.value || '').trim().toLowerCase();
    const { headers, rows } = data;

    // 동명이인 계산
    const nameCol  = headers.findIndex(h => ['성명','이름','학생이름'].some(k => h.includes(k)));
    const gradeCol = headers.findIndex(h => h.includes('학년'));
    const dupSet   = new Set();
    if (_dupOnly && nameCol >= 0) {
      const cnt = {};
      rows.forEach((r, i) => {
        const nm    = (r[nameCol] || '').replace(/[A-Z]+$/, '').trim();
        const grade = gradeCol >= 0 ? (r[gradeCol] || '') : '';
        const key   = `${grade}||${nm}`;
        if (nm) cnt[key] = (cnt[key] || []).concat(i);
      });
      Object.values(cnt).forEach(idxs => {
        if (idxs.length >= 2) idxs.forEach(i => dupSet.add(i));
      });
    }

    // 비고 컬럼 탐색 (보류 색상용)
    const noteCol = (() => {
      const candidates = headers.reduce((a, h, i) => (h.includes('비고') || h.includes('사유')) ? [...a, i] : a, []);
      return candidates.length ? candidates[candidates.length - 1] : headers.length - 1;
    })();

    const filtered = rows.reduce((acc, row, i) => {
      if (row.every(v => !String(v).trim())) return acc;  // 빈 행 제외
      if (keyword && !row.join(' ').toLowerCase().includes(keyword)) return acc;
      if (_dupOnly && !dupSet.has(i)) return acc;
      acc.push({ row, i, noteVal: row[noteCol] || '' });
      return acc;
    }, []);

    const table = _el('run-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = '<tr>' + headers.map(h => `<th>${_esc(h)}</th>`).join('') + '</tr>';
    tbody.innerHTML = filtered.map(({ row, i, noteVal }) => {
      const isHold     = noteVal.includes('보류:') && !noteVal.includes('자동 제외');
      const isAutoSkip = noteVal.includes('자동 제외');
      const isDup      = dupSet.has(i);
      const cls = isHold ? 'row-hold' : isAutoSkip ? 'row-skip' : isDup ? 'row-dup' : '';
      return `<tr class="${cls}">${row.map(v => `<td>${_esc(v)}</td>`).join('')}</tr>`;
    }).join('');

    _el('run-preview-info').textContent =
      `시트: ${_currentSheet} | 행 수: ${filtered.length}`;
  }

  function filterTable() { _renderRunTable(); }

  function toggleDup() {
    _dupOnly = !_dupOnly;
    _el('btn-run-dup')?.classList.toggle('active', _dupOnly);
    _renderRunTable();
  }

  // ──────────────────────────────────────────────
  // 파일 / 폴더 열기
  // ──────────────────────────────────────────────
  async function openFile(path) {
    if (!path) return;
    await bridge.openFile(path);
  }

  async function openFolder() {
    const path = _outputFiles[0]?.path || state.work_root;
    if (!path) { toast('열 폴더가 없습니다. 먼저 작업을 실행해 주세요.', 'warn'); return; }
    await bridge.openFolder(path);
  }

  // ──────────────────────────────────────────────
  // 로그
  // ──────────────────────────────────────────────
  function showLog() { showLogDialog('실행 로그', state.last_run_logs); }

  // ──────────────────────────────────────────────
  // 초기화
  // ──────────────────────────────────────────────
  function reset() {
    _outputFiles   = [];
    _currentFile   = null;
    _sheetData     = {};
    _currentSheet  = null;
    _pendingSheets = [];
    _dupOnly       = false;
    _lastRunData   = null;

    _setBadge('idle', '실행 전');
    _el('run-info').textContent        = '먼저 스캔을 통과해야 실행할 수 있습니다.';
    _el('run-hold-warn').style.display = 'none';
    _el('btn-goto-notice').style.display = 'none';
    _el('btn-run').disabled             = true;
    _el('output-file-list').innerHTML   = '<span class="muted">실행 전</span>';

    const tabs = _el('sheet-tabs');
    if (tabs) tabs.innerHTML = '';
    const table = _el('run-table');
    if (table) {
      table.querySelector('thead').innerHTML = '';
      table.querySelector('tbody').innerHTML = '';
    }
    _el('run-preview-info').textContent = '시트: - | 행 수: -';

    ['sum-school','sum-year','sum-freshmen','sum-teacher',
     'sum-transfer','sum-withdraw','sum-transfer-check','sum-withdraw-check']
      .forEach(id => { const el = _el(id); if (el) el.textContent = '-'; });
  }

  function _setBadge(type, text) {
    const el = _el('run-status-badge');
    if (!el) return;
    el.className   = `status-badge badge-${type}`;
    el.textContent = text;
  }

  function _esc(str) {
    return String(str ?? '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  return {
    start, onFinished, onFailed, onPreviewLoaded,
    filterTable, toggleDup, openFile, openFolder,
    showLog, reset,
  };

})();


// ─────────────────────────────────────────────────
// Diff 탭
// ─────────────────────────────────────────────────
const Diff = (() => {

  // ──────────────────────────────────────────────
  // 실행 시작
  // ──────────────────────────────────────────────
  async function start() {
    if (state.isDiffRunning) return;
    if (!state.work_root)       { toast('작업 폴더가 설정되지 않았습니다.', 'warn'); return; }
    if (!state.selected_school) { toast('학교를 먼저 선택해 주세요.', 'warn'); return; }

    const targetYear = parseInt(_el('diff-target-year')?.value || new Date().getFullYear(), 10);

    state.isDiffRunning = true;
    _el('btn-run-diff').disabled = true;
    _setBadge('running', '실행 중');
    _el('diff-result-info').textContent = '실행 중...';

    const params = {
      work_root:         state.work_root,
      school_name:       state.selected_school,
      target_year:       targetYear,
      school_start_date: state.school_start_date,
      work_date:         state.work_date,
      roster_xlsx:       state.roster_log_path || '',
      col_map:           state.roster_col_map  || {},
    };

    const res = JSON.parse(await bridge.startRunDiff(JSON.stringify(params)));
    if (!res.ok) {
      state.isDiffRunning = false;
      _el('btn-run-diff').disabled = false;
      _setBadge('err', '실행 실패');
      _el('diff-result-info').textContent = res.error || '명단 비교 시작 실패';
    }
  }

  // ──────────────────────────────────────────────
  // Diff Scan 완료 (사전 점검 결과 — 현재 bridge에서 별도 시그널 없음)
  // ──────────────────────────────────────────────
  function onScanFinished(data) {
    // 향후 사전 점검 결과 UI 표시 시 사용
  }

  function onScanFailed(error) {
    _el('diff-result-info').textContent = `사전 점검 실패: ${error}`;
  }

  // ──────────────────────────────────────────────
  // Diff 실행 완료 (main.js에서 호출)
  // ──────────────────────────────────────────────
  function onFinished(data) {
    _el('btn-run-diff').disabled = false;

    if (!data.ok) {
      _setBadge('err', '실행 실패');
      const err = (data.logs || []).find(l => l.level === 'error');
      _el('diff-result-info').textContent = err ? err.message : '실행 실패';
      return;
    }

    _setBadge('ok', '완료');

    const lines = [
      `일치 (정상 재학생): ${data.matched_count}명`,
      `학교 명단에만 있음: ${data.compare_only_count}명`,
      `명부에만 있음: ${data.roster_only_count}명`,
      `판정 불가: ${data.unresolved_count}명`,
      `──`,
      `자동 분류 전입 ${data.transfer_in_done}명 / 확인 필요 ${data.transfer_in_hold}명`,
      `자동 분류 전출 ${data.transfer_out_done}명 / 확인 필요 ${data.transfer_out_hold}명`,
    ];
    _el('diff-result-info').textContent = lines.join('\n');

    // 출력 파일 목록
    const filesEl = _el('diff-output-files');
    if (filesEl) {
      filesEl.innerHTML = (data.output_files || []).map(f => `
        <div class="output-file-item">
          <span class="output-file-name" onclick="Run.openFile('${f.path}')">${f.name}</span>
        </div>
      `).join('') || '<span class="muted">생성된 파일 없음</span>';
    }
  }

  function onFailed(error) {
    _el('btn-run-diff').disabled = false;
    _setBadge('err', '실행 실패');
    _el('diff-result-info').textContent = `예기치 못한 오류: ${error}`;
    toast('명단 비교 오류: ' + error, 'err');
  }

  // ──────────────────────────────────────────────
  // 로그
  // ──────────────────────────────────────────────
  function showLog() { showLogDialog('명단 비교 로그', state.last_diff_logs); }

  // ──────────────────────────────────────────────
  // 초기화
  // ──────────────────────────────────────────────
  function reset() {
    _setBadge('idle', '실행 전');
    _el('diff-result-info').textContent = '실행 후 결과가 표시됩니다.';
    const filesEl = _el('diff-output-files');
    if (filesEl) filesEl.innerHTML = '';
    _el('btn-run-diff').disabled = true;
    // 연도는 올해로 복원
    const yearEl = _el('diff-target-year');
    if (yearEl) yearEl.value = new Date().getFullYear();
  }

  function _setBadge(type, text) {
    const el = _el('diff-status-badge');
    if (!el) return;
    el.className   = `status-badge badge-${type}`;
    el.textContent = text;
  }

  return { start, onScanFinished, onScanFailed, onFinished, onFailed, showLog, reset };

})();
