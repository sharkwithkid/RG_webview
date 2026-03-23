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
 *           btn-blank-only, btn-issue-only, btn-dup-only,
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
  let _filterState  = { blank: false, issue: false, dup: false };

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
      const err = (data.logs || []).find(l => l.level === 'error');
      _setBadge('err', '스캔 실패');
      _setMessage(err ? err.message : '스캔 중 오류가 발생했습니다.');
      _hideSchoolKindWarn();
      App.setStepState(2, 'warn');
      return;
    }

    // 스텝 상태
    App.setStepState(2, 'done');

    // 학교 구분 판별 실패 여부
    const kindWarn = (data.logs || []).some(l =>
      l.level === 'warn' && l.message.includes('학교 구분을 자동으로 판별하지 못했습니다')
    );
    _setSchoolKindWarn(kindWarn);

    // 경고 / 완료 뱃지
    const warnLogs = (data.logs || []).filter(l => l.level === 'warn');
    if (warnLogs.length) {
      _setBadge('warn', '경고');
      if (data.need_roster && !data.roster_path) {
        _setMessage('학생명부를 찾지 못했습니다. 학년도 아이디 규칙을 직접 입력하거나 명부를 추가한 뒤 재스캔해 주세요.');
      } else {
        _setMessage(`경고 ${warnLogs.length}건 — ${warnLogs[0].message}`);
      }
    } else {
      _setBadge('ok', '스캔 완료');
      _setMessage('스캔 완료 — 이상 없음');
    }

    // 학년도 아이디 규칙 갱신
    _updateGradeMap(data);

    // 스캔 표 갱신
    _applyScanTable(data.items || []);

    // 뷰어 자동 펼침 + 첫 행 미리보기 요청
    _openViewer();
    _loadFirstPreview(data.items || []);

    // 실행 버튼 활성
    if (data.can_execute) {
      _el('btn-run').disabled = false;
      _el('btn-goto-run').style.display = '';
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
      const row = KIND_ROW[item.kind];
      if (row === undefined) return;
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${item.kind}"]`);
      if (!tr) return;

      const cells = tr.querySelectorAll('td');
      // cells: [0]구분 [1]파일명 [2]시트 [3]자동감지 [4]수정시작행 [5]확인
      if (cells[1]) {
        cells[1].className = 'file-link';
        cells[1].textContent = item.file_name || '';
        cells[1].title = '클릭: 뷰어로 보기 · 더블클릭: 파일 열기';
        cells[1].onclick = () => {
          // 기존 active 행 해제 후 현재 행 활성화
          document.querySelectorAll('#scan-tbody tr.viewer-active')
            .forEach(r => r.classList.remove('viewer-active'));
          tr.classList.add('viewer-active');
          _requestPreview(item.kind);
        };
        cells[1].ondblclick = () => { if (item.file_path) bridge.openFile(item.file_path); };
      }
      if (cells[2]) cells[2].textContent = item.sheet_name  || '';
      if (cells[3]) cells[3].textContent = item.header_row != null ? String(item.header_row) : '';

      // 수정 시작행 스핀 초기값 세팅
      const spinVal = _el(`spin-${item.kind}`);
      if (spinVal && item.data_start_row != null) {
        spinVal.textContent = String(item.data_start_row);
        spinVal.style.color = '#0F172A';
      }
    });

    // 파일 없는 구분 비활성화
    Object.keys(KIND_ROW).forEach(kind => {
      if (!presentKinds.has(kind)) {
        const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
        if (!tr) return;
        const cells = tr.querySelectorAll('td');
        if (cells[1]) { cells[1].className = ''; cells[1].textContent = ''; cells[1].onclick = null; }
        if (cells[2]) cells[2].textContent = '';
        if (cells[3]) cells[3].textContent = '';
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
    if (state.isPreviewLoading) return;

    const cached = _previewData[kind];
    if (cached && cached.rows) {
      // 이미 로드된 경우 바로 렌더
      _currentKind = kind;
      _renderPreview(kind);
      return;
    }

    // 스캔 데이터에서 파일 메타 찾기
    const item = (_lastScanData?.items || []).find(i => i.kind === kind);
    if (!item || !item.file_path) {
      _el('preview-warn').textContent = `${kind} 파일 경로 정보가 없습니다.`;
      return;
    }

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
    _currentKind = kind;
    _renderPreview(kind);
  }

  function onPreviewFailed(kind, error) {
    _el('preview-warn').textContent = `미리보기 실패 (${kind}): ${error}`;
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
      ` | 총 ${data.total_count ?? '?'}행` +
      (data.truncated ? ` (${data.rows.length}행까지 표시)` : '');

    _el('preview-warn').textContent =
      data.truncated ? `※ 전체 ${data.total_count}행 중 ${data.rows.length}행만 표시됩니다.` : '';

    _renderTable(data);
  }

  function _renderTable(data) {
    const keyword   = (_el('preview-search')?.value || '').trim().toLowerCase();
    const blankOnly = _filterState.blank;
    const issueOnly = _filterState.issue;
    const dupOnly   = _filterState.dup;

    const columns   = data.columns || [];
    const rows      = data.rows    || [];
    const issueSet  = new Set(data.issue_rows || []);

    // 동명이인 계산
    const nameCol  = columns.findIndex(h => ['성명','이름','학생이름'].some(k => h.includes(k)));
    const gradeCol = columns.findIndex(h => h.includes('학년'));
    const dupSet   = new Set();
    if (dupOnly && nameCol >= 0) {
      const cnt = {};
      rows.forEach((r, i) => {
        const nm    = (r[nameCol] || '').replace(/[A-Z]+$/, '').trim();
        const grade = gradeCol >= 0 ? (r[gradeCol] || '') : '';
        const key   = `${grade}||${nm}`;
        if (nm) cnt[key] = (cnt[key] || []).concat(i);
      });
      Object.values(cnt).forEach(idxs => { if (idxs.length >= 2) idxs.forEach(i => dupSet.add(i)); });
    }

    // 필터
    const filtered = rows.reduce((acc, row, i) => {
      const rowText = row.join(' ').toLowerCase();
      if (keyword && !rowText.includes(keyword)) return acc;
      const isBlank = row.every(v => !v.trim());
      const isIssue = issueSet.has(i);
      const isDup   = dupSet.has(i);
      if (blankOnly && !isBlank) return acc;
      if (issueOnly && !isIssue) return acc;
      if (dupOnly   && !isDup)   return acc;
      acc.push({ row, i, isIssue, isDup });
      return acc;
    }, []);

    // 렌더
    const table = _el('preview-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = '<tr>' + columns.map(h => `<th>${_esc(h)}</th>`).join('') + '</tr>';
    tbody.innerHTML = filtered.map(({ row, i, isIssue, isDup }) => {
      const cls = isIssue ? 'row-hold' : isDup ? 'row-dup' : '';
      return `<tr class="${cls}">${row.map(v => `<td>${_esc(v)}</td>`).join('')}</tr>`;
    }).join('');
  }

  // ──────────────────────────────────────────────
  // 필터 토글
  // ──────────────────────────────────────────────
  function toggleFilter(key) {
    _filterState[key] = !_filterState[key];
    const btnId = { blank: 'btn-blank-only', issue: 'btn-issue-only', dup: 'btn-dup-only' }[key];
    _el(btnId)?.classList.toggle('active', _filterState[key]);
    if (_currentKind) _renderTable(_previewData[_currentKind] || {});
  }

  function filterPreview() {
    if (_currentKind) _renderTable(_previewData[_currentKind] || {});
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

  // ──────────────────────────────────────────────
  // 초기화 (학교 변경 시)
  // ──────────────────────────────────────────────
  function reset() {
    _lastScanData  = null;
    _previewData   = {};
    _currentKind   = null;
    _filterState   = { blank: false, issue: false, dup: false };

    _setBadge('idle', '스캔 전');
    _setMessage('파일 내용 스캔을 실행해 주세요.');
    _hideSchoolKindWarn();

    // 스캔 표 초기화
    Object.keys(KIND_ROW).forEach(kind => {
      const tr = document.querySelector(`#scan-tbody tr[data-kind="${kind}"]`);
      if (!tr) return;
      const cells = tr.querySelectorAll('td');
      [1, 2, 3].forEach(c => { if (cells[c]) cells[c].textContent = ''; });
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

    _el('btn-goto-run').style.display = 'none';
    _el('btn-run').disabled            = true;

    // 필터 버튼 초기화
    ['btn-blank-only','btn-issue-only','btn-dup-only'].forEach(id => _el(id)?.classList.remove('active'));
  }

  // 실행 탭으로 이동 (확인 체크박스 검증)
  function goToRun() {
    const lastData = _lastScanData;
    if (!lastData) return;

    // 스캔된 종류 중 체크 안 된 것 찾기
    const presentKinds = (lastData.items || []).map(i => i.kind);
    const unchecked = presentKinds.filter(kind => {
      const chk = _el(`chk-${kind}`);
      return chk && !chk.checked;
    });

    if (unchecked.length) {
      toast(`검수 확인이 필요합니다: ${unchecked.join(', ')}`, 'warn', 4000);
      return;
    }

    App.goStep(3);
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
  };

})();
