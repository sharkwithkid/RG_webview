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
  let _previewMeta  = {}; // { sheetName: { actual_count, max_row, displayed_count } }
  let _noticeDupRows = new Set();         // 학생 안내문 시트 동명이인 행 번호
  let _noticeTeacherDupRows = new Set(); // 교사 안내문 시트 동명이인 행 번호

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
    _el('run-info').textContent = '작업을 실행하는 중입니다.';
    _el('run-hold-warn').style.display = 'none';

    const res = JSON.parse(await bridge.startRunMain(JSON.stringify(params)));
    if (!res.ok) {
      state.isRunning = false;
      _el('btn-run').disabled = false;
      _setBadge('err', '실패');
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
    _noticeDupRows        = new Set(data.notice_dup_rows || []);
    _noticeTeacherDupRows = new Set(data.notice_teacher_dup_rows || []);

    if (!data.ok) {
      const status = data.status || null;
      const errMsg = (status?.messages || []).find(m => m.level === 'error')?.text
        || '작업 실행 중 오류가 발생했습니다.';
      _setBadge('err', '실패');
      _el('run-info').textContent = '실행 로그와 결과 카드를 확인해 주세요.';
      _renderRunStatusCard(status || { level: 'error', summary_text: errMsg, detail_messages: [errMsg] }, data);
      App.setStepState(3, 'warn');
      // 토스트 제거 — 카드가 남아있어서 중복
      return;
    }

    // 스텝 상태
    App.setStepState(3, 'done');
    App.setStepState(4, 'active');

    // 뱃지 + 카드 — status 하나만 본다 (logs 파싱 없음)
    const holdWarn = _el('run-hold-warn');
    const status = data.status || null;
    const level  = status?.level || 'ok';

    StatusUI.renderBadge('run-status-badge', status?.badge, '완료');

    if (level === 'error') {
      _el('run-info').textContent = '실행 로그와 결과 카드를 확인해 주세요.';
      _renderRunStatusCard(status, data);
      App.setStepState(3, 'warn');
    } else if (level === 'warn' || level === 'hold') {
      _el('run-info').textContent = '실행 로그와 결과 카드를 확인해 주세요.';
      _renderRunStatusCard(status, data);
      App.setStepState(3, 'warn');
    } else {
      _el('run-info').textContent = '작업이 완료되었습니다.';
      holdWarn.style.display = 'none';
      holdWarn.innerHTML = '';
      App.setStepState(3, 'done');
    }

    // 상태 요약 그리드
    _updateSummary(data);

    // 출력 파일 목록
    _outputFiles = data.output_files || [];
    _renderOutputFiles();

    // 내부 작업 이력 자동 저장 (실행 완료 기준)
    (async () => {
      try {
        const scanData = Scan.getLastScanData ? Scan.getLastScanData() : null;
        const items = scanData?.items || [];

        const freshmenItem    = items.find(i => i.kind === '신입생');
        const transferInItem  = items.find(i => i.kind === '전입생');
        const transferOutItem = items.find(i => i.kind === '전출생');
        const teacherItem     = items.find(i => i.kind === '교직원');

        // 전입/전출은 run payload의 done 값 우선 사용
        // scan row_count는 보류 포함 원본 건수라 실제 처리 수와 다를 수 있음
        const counts = {
          '신입생': data.freshmen_count    ?? freshmenItem?.row_count    ?? 0,
          '전입생': data.transfer_in_done  ?? transferInItem?.row_count  ?? 0,
          '전출생': data.transfer_out_done ?? transferOutItem?.row_count ?? 0,
          '교직원': data.teacher_count     ?? teacherItem?.row_count     ?? 0,
        };

        const entry = {
          last_date: state.work_date || _todayStr(),
          worker: state.worker_name || '',
          counts,
          run_completed: true,
          master_recorded: false,
        };

        const schoolYear = (state.work_date || _todayStr()).slice(0, 4);
        const histRes = JSON.parse(
          await bridge.saveWorkHistory(
            schoolYear,
            state.selected_school,
            JSON.stringify(entry)
          )
        );

        if (!histRes.ok) {
          console.error('[HISTORY][AUTO] 저장 실패:', histRes.error);
          return;
        }

        const SHORT = { '신입생': '신입', '전입생': '전입', '전출생': '전출', '교직원': '교직' };
        const countStr = Object.entries(counts)
          .filter(([, v]) => v)
          .map(([k, v]) => `${SHORT[k] ?? k} ${v}`)
          .join(' · ');

        let histText = `마지막 작업 · ${entry.last_date}`;
        if (entry.worker) histText += ` (${entry.worker})`;
        if (countStr) histText += `\n${countStr}`;
        histText += `\n전체 명단 반영 전`;

        Panel.updateSchoolInfo({
          school_name: _el('current-school')?.textContent || state.selected_school,
          history_text: histText,
        });
      } catch (e) {
        console.error('[HISTORY][AUTO] 예외:', e);
      }
    })();

    // 명단 기록 버튼 활성
    // record: roster_log_path 있을 때만 / open: 학교 선택됐으면 항상
    state.pending_roster_log = !!state.roster_log_path;
    Panel.setRosterBtns(!!state.roster_log_path, !!state.selected_school);

    // 안내문 탭으로 플로팅 버튼
    App.setFloatingNext(true, 'notice');
  }


  // _collectStatusMessages 제거 — status.messages 직접 사용

  function onFailed(error) {
    _el('btn-run').disabled = false;
    _setBadge('err', '실패');
    _el('run-info').textContent = '실행 로그와 결과 카드를 확인해 주세요.';
    _renderRunStatusCard({
      level: 'error',
      summary_text: '실행 중 오류가 발생했습니다.',
      detail_messages: [String(error?.message || error || '예기치 못한 오류가 발생했습니다.')],
    });
    App.setStepState(3, 'warn');
    // 토스트 제거 — 카드가 있어서 중복
  }


  function _renderRunStatusCard(status, data=null) {
    const holdWarn = _el('run-hold-warn');
    if (!holdWarn) return;

    const level = status?.level === 'error' ? 'error' : 'warn';
    const details = Array.from(new Set(
      (status?.messages || status?.detail_messages || [])
        .map(m => String(m?.text || m || '').trim())
        .filter(Boolean)
    ));
    const summary = String(status?.summary_text || '').trim();
    const action = String(status?.action_text || '').trim();

    if (!summary && !details.length && !action) {
      holdWarn.style.display = 'none';
      holdWarn.innerHTML = '';
      return;
    }

    holdWarn.classList.toggle('error', level === 'error');
    const html = (typeof StatusUI !== 'undefined' && StatusUI.normalizeStatusCard)
      ? StatusUI.normalizeStatusCard(details, level, { summary_text: summary, action_text: action })
      : null;

    holdWarn.innerHTML = html || '';
    holdWarn.style.display = 'block';
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
    _previewMeta   = {};
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
    const { sheet_name, columns, rows, row_colors } = payload;
    _sheetData[sheet_name] = { headers: columns, rows, rowColors: row_colors || [] };
    _previewMeta[sheet_name] = {
      actual_count: Number.isFinite(payload.actual_count) ? payload.actual_count : rows.length,
      max_row: Number.isFinite(payload.max_row) ? payload.max_row : null,
      displayed_count: Number.isFinite(payload.displayed_count) ? payload.displayed_count : rows.length,
    };

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
      // state.isPreviewLoading은 bridge._is_previewing과 별개로 관리
      setTimeout(() => {
        bridge.startPreview(JSON.stringify({
          kind:           'run_output',
          file_path:      _currentFile,
          sheet_name:     next,
          header_row:     1,
          data_start_row: 2,
        }));
      }, 50);
    } else {
      state.isPreviewLoading = false;
    }
  }

  function _switchSheet(name) {
    _currentSheet = name;
    document.querySelectorAll('.sheet-tab').forEach(b =>
      b.classList.toggle('active', b.dataset.sheet === name)
    );

    // 시트 전환 시 페이지 스크롤 튐 방지:
    // 전환 전 tbody 높이를 min-height로 고정 → 렌더 후 해제
    const runTable = _el('run-table');
    const wrap = runTable?.closest('.preview-table-wrap');
    if (wrap) {
      wrap.style.minHeight = wrap.offsetHeight + 'px';
    }

    _renderRunTable();

    if (wrap) {
      // 렌더 완료 후 min-height 해제 (다음 프레임에서)
      requestAnimationFrame(() => { wrap.style.minHeight = ''; });
      wrap.scrollTop = 0;
    }
  }

  function _renderRunTable() {
    const data = _sheetData[_currentSheet];
    if (!data) return;

    const keyword = (_el('run-search')?.value || '').trim().toLowerCase();
    const { headers, rows, rowColors } = data;

    // 동명이인: Python 코어가 판별한 running_no(1-based) 기준
    // _noticeDupRows는 안내문 시트 기준 행 번호 → No. 컬럼(첫 번째 열) 값과 매칭
    const noCol = headers.findIndex(h => h === 'No.' || h === 'No' || h === 'NO' || h === 'no');

    // 비고 컬럼 탐색 (보류 색상용)
    const noteCol = (() => {
      const candidates = headers.reduce((a, h, i) => (h.includes('비고') || h.includes('사유')) ? [...a, i] : a, []);
      return candidates.length ? candidates[candidates.length - 1] : headers.length - 1;
    })();

    const filtered = rows.reduce((acc, row, i) => {
      if (row.every(v => !String(v).trim())) return acc;  // 빈 행 제외
      if (keyword && !row.join(' ').toLowerCase().includes(keyword)) return acc;
      // 동명이인 필터: No. 컬럼 값이 _noticeDupRows에 있는 행만
      const rowNo = noCol >= 0 ? parseInt(row[noCol], 10) : NaN;
      // 동명이인: 시트 종류에 따라 해당 Set 적용
      const sheetName = _currentSheet || '';
      const isStudentSheet = sheetName.includes('학생');
      const isTeacherSheet = sheetName.includes('선생') || sheetName.includes('교사') || sheetName.includes('teacher');
      const isDupRow = (!isNaN(rowNo)) && (
        (isStudentSheet && _noticeDupRows.size > 0 && _noticeDupRows.has(rowNo)) ||
        (isTeacherSheet && _noticeTeacherDupRows.size > 0 && _noticeTeacherDupRows.has(rowNo))
      );
      if (_dupOnly && !isDupRow) return acc;
      acc.push({ row, i, noteVal: row[noteCol] || '', rowColor: (rowColors && rowColors[i]) || null, isDupRow });
      return acc;
    }, []);

    const table = _el('run-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = '<tr>' + headers.map(h => `<th>${_esc(h)}</th>`).join('') + '</tr>';
    tbody.innerHTML = filtered.map(({ row, i, noteVal, rowColor, isDupRow }) => {
      const isHold     = noteVal.includes('보류:') && !noteVal.includes('자동 제외');
      const isAutoSkip = noteVal.includes('자동 제외');
      const cls = isHold ? 'row-hold' : isAutoSkip ? 'row-skip' : isDupRow ? 'row-dup' : '';
      return `<tr class="${cls}">${row.map(v => `<td>${_esc(v)}</td>`).join('')}</tr>`;
    }).join('');

    const meta = _previewMeta[_currentSheet] || {};
    const totalCount = Number.isFinite(meta.actual_count) ? meta.actual_count : filtered.length;
    const maxRow = Number.isFinite(meta.max_row) ? meta.max_row : null;
    _el('run-preview-info').textContent =
      `시트: ${_currentSheet} | 실제 ${totalCount}행${maxRow != null ? ` | 최대 ${maxRow}행` : ''}`;
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
    _previewMeta    = {};
    _noticeDupRows        = new Set();
    _noticeTeacherDupRows = new Set();

    _setBadge('idle', '대기');
    _el('run-info').textContent = '스캔을 통과한 후 작업을 실행하고 결과 파일을 확인합니다.';
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

  function _handleDiffRosterDateMismatch(basisDate) {
    const workDate = state.work_date;

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
          <label class="confirm-modal-option selected">
            <input type="radio" name="diff-cm-date" value="basis" checked>
            <div>
              <div class="confirm-modal-option-label">수정일 사용 — ${basisDate}</div>
              <div class="confirm-modal-option-desc">명부 파일의 마지막 수정일을 기준으로 합니다.</div>
            </div>
          </label>
          <label class="confirm-modal-option">
            <input type="radio" name="diff-cm-date" value="work">
            <div>
              <div class="confirm-modal-option-label">작업일 사용 — ${workDate}</div>
              <div class="confirm-modal-option-desc">오늘 작업일을 기준으로 진행합니다.</div>
            </div>
          </label>
        </div>
        <div class="confirm-modal-footer">
          <button class="btn-primary" id="diff-cm-date-confirm" style="height:36px;padding:0 20px">확인</button>
        </div>
      </div>`;

    document.body.appendChild(backdrop);

    backdrop.querySelectorAll('input[type="radio"]').forEach(radio => {
      radio.addEventListener('change', () => {
        backdrop.querySelectorAll('.confirm-modal-option').forEach(opt => opt.classList.remove('selected'));
        radio.closest('.confirm-modal-option').classList.add('selected');
      });
    });

    backdrop.querySelector('#diff-cm-date-confirm').addEventListener('click', () => {
      const selected = backdrop.querySelector('input[name="diff-cm-date"]:checked')?.value;
      backdrop.remove();
      _runWithBasisDate(selected === 'work' ? workDate : basisDate);
    });
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
