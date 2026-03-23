/**
 * status_panel.js — StatusPanel 로직
 *
 * 의존: main.js (state, bridge, App, _el, _todayStr)
 * HTML ID: school-input, school-dropdown, btn-apply-school,
 *           current-school, school-history, school-status, last-work,
 *           btn-grade-toggle, grade-map-body, grade-rows,
 *           chk-arrived, arrived-date, chk-sent, sent-date,
 *           btn-record-roster, btn-open-roster,
 *           btn-run-diff (diff 버튼 활성/비활성)
 */

'use strict';

const Panel = (() => {

  let _searchTimer  = null;
  let _hlIndex      = -1;    // 드롭다운 키보드 탐색 인덱스
  let _gradeOpen    = false;
  let _maxGrade     = 6;

  // ──────────────────────────────────────────────
  // 초기화 (main.js에서 학교명 로드 후 호출)
  // ──────────────────────────────────────────────
  function init(names) {
    state.school_names = names || [];
    _buildGradeRows();
    _setStatus(
      state.school_names.length
        ? `학교명 검색 준비 완료 · ${state.school_names.length}개`
        : '학교 목록 0건: 명단 파일 및 열 매핑을 확인하세요.',
      state.school_names.length ? '' : 'warn'
    );
  }

  // 작업 컨텍스트 세팅 (SetupPage 완료 후 main.js가 호출)
  function setWorkContext({ work_date, arrived_date }) {
    const d = work_date || _todayStr();
    // 도착일: 저장된 값 우선, 없으면 work_date
    _setDateInput('arrived-date', arrived_date || d);
    _setDateInput('sent-date',    d);
    // DatePicker 트리거 텍스트 동기화
    if (typeof DatePicker !== 'undefined') {
      DatePicker.setValue('arrived-date', arrived_date || d);
      DatePicker.setValue('sent-date',    d);
    }
  }

  // ──────────────────────────────────────────────
  // 학교 검색 입력
  // ──────────────────────────────────────────────
  function onInput(text) {
    clearTimeout(_searchTimer);
    _hlIndex = -1;
    _openDropdown();
    _searchTimer = setTimeout(() => _applySearch(text.trim()), 150);
  }

  function _applySearch(keyword) {
    const names = state.school_names;
    if (!names.length) {
      _closeDropdown();
      _setStatus('학교 목록을 불러오지 못했습니다.', 'warn');
      _setApplyBtn(false);
      return;
    }

    let matched;
    if (!keyword) {
      matched = names.slice(0, 100);
      _setStatus(`전체 학교 목록 ${names.length}개`);
      _setApplyBtn(false);
    } else {
      matched = names
        .filter(n => n.toLowerCase().includes(keyword.toLowerCase()))
        .slice(0, 100);
      if (matched.length) {
        _setStatus(`검색 결과 ${matched.length}건`);
      } else {
        _setStatus('DB에 일치하는 학교가 없습니다.', 'warn');
        _closeDropdown();
      }
      _setApplyBtn(_isExact(keyword));
    }

    _renderDropdown(matched);
  }

  function _renderDropdown(names) {
    const dd = _el('school-dropdown');
    dd.innerHTML = '';
    names.forEach(name => {
      const el = document.createElement('div');
      el.className   = 'school-item';
      el.textContent = name;
      el.onclick     = () => _selectName(name);
      dd.appendChild(el);
    });
  }

  function _selectName(name) {
    _el('school-input').value = name;
    _closeDropdown();
    const exact = _isExact(name);
    _setApplyBtn(exact);
    _setStatus(exact ? '적용 가능한 학교입니다.' : 'DB에 없는 학교입니다.', exact ? 'ok' : 'warn');
  }

  // ──────────────────────────────────────────────
  // 키보드 탐색
  // ──────────────────────────────────────────────
  function onKeyDown(e) {
    const dd    = _el('school-dropdown');
    const items = dd.querySelectorAll('.school-item');

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      _hlIndex = Math.min(_hlIndex + 1, items.length - 1);
      _highlight(items);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      _hlIndex = Math.max(_hlIndex - 1, 0);
      _highlight(items);
    } else if (e.key === 'Enter') {
      if (_hlIndex >= 0 && items[_hlIndex]) {
        _selectName(items[_hlIndex].textContent);
      } else {
        apply();
      }
    } else if (e.key === 'Escape') {
      _closeDropdown();
    }
  }

  function _highlight(items) {
    items.forEach((el, i) => el.classList.toggle('hl', i === _hlIndex));
    if (items[_hlIndex]) {
      items[_hlIndex].scrollIntoView({ block: 'nearest' });
      // 키보드 탐색 중에만 input 값 교체 (타이핑 중 커서 점프 방지)
      const inp = _el('school-input');
      const name = items[_hlIndex].textContent;
      inp.value = name;
      // 커서를 끝으로 이동
      inp.setSelectionRange(name.length, name.length);
      _setApplyBtn(_isExact(name));
    }
  }

  // ──────────────────────────────────────────────
  // 학교 적용
  // ──────────────────────────────────────────────
  async function apply() {
    const name = (_el('school-input').value || '').trim();
    if (!_isExact(name)) return;

    _setApplyBtn(false);
    _closeDropdown();

    // main.js → App.onSchoolSelected
    await App.onSchoolSelected(name);
  }

  // 학교 선택 후 UI 갱신 (main.js → App.onSchoolSelected 내부에서 호출)
  function updateSchoolInfo({ school_name, history_text, last_work_text }) {
    _el('current-school').textContent = school_name || '-';

    // 선택 완료 후 status 메시지 클리어
    _setStatus('');

    const hist = _el('school-history');
    if (history_text) { hist.textContent = history_text; hist.style.display = ''; }
    else               hist.style.display = 'none';

    const last = _el('last-work');
    if (last_work_text) { last.textContent = last_work_text; last.style.display = ''; }
    else                 last.style.display = 'none';
  }

  // ──────────────────────────────────────────────
  // 초기화 (새 학교 시작 / 학교 선택 리셋)
  // ──────────────────────────────────────────────
  function reset() {
    _el('school-input').value     = '';
    _el('current-school').textContent = '-';
    _el('school-history').style.display = 'none';
    _el('last-work').style.display      = 'none';
    _setStatus('학교명을 입력해 검색하세요.');
    _setApplyBtn(false);
    _closeDropdown();

    _el('chk-arrived').checked = false;
    _el('chk-sent').checked    = false;
    _el('btn-record-roster').disabled = true;
    _el('btn-open-roster').disabled   = true;
  }

  // ──────────────────────────────────────────────
  // 새 학교 시작 버튼
  // ──────────────────────────────────────────────
  function newSchool() {
    if (state.pending_roster_log) {
      if (!confirm('명단이 아직 기록되지 않았습니다.\n그래도 새 학교 작업을 시작하시겠습니까?')) return;
    } else {
      if (!confirm('새 학교 작업을 시작하시겠습니까?\n현재 작업 내용이 초기화됩니다.')) return;
    }
    App.resetToSchoolSelect();
  }

  // ──────────────────────────────────────────────
  // 도착일 / 발송일 정보 읽기 (main.js / run_tab.js 에서 호출)
  // ──────────────────────────────────────────────
  function getArrivedInfo() {
    return {
      checked: _el('chk-arrived').checked,
      date:    _el('arrived-date').value || '',
    };
  }

  function getSentInfo() {
    return {
      checked: _el('chk-sent').checked,
      date:    _el('sent-date').value || '',
    };
  }

  // ──────────────────────────────────────────────
  // 명단 기록 / 명단 파일 열기 버튼 활성화
  // ──────────────────────────────────────────────
  function setRosterBtns(recordEnabled, openEnabled) {
    _el('btn-record-roster').disabled = !recordEnabled;
    _el('btn-open-roster').disabled   = !openEnabled;
  }

  async function recordRoster() {
    if (!state.roster_log_path || !state.selected_school) return;

    const arrived = getArrivedInfo();
    const sent    = getSentInfo();
    const scanData = Scan.getLastScanData();
    const kindFlags = scanData ? {
      신입생: !!(scanData.items || []).find(i => i.kind === '신입생'),
      전입생: !!(scanData.items || []).find(i => i.kind === '전입생'),
      전출생: !!(scanData.items || []).find(i => i.kind === '전출생'),
      교직원: !!(scanData.items || []).find(i => i.kind === '교직원'),
    } : {};

    const params = {
      xlsx_path:          state.roster_log_path,
      school_name:        state.selected_school,
      worker:             state.worker_name,
      kind_flags:         kindFlags,
      email_arrived_date: arrived.checked ? arrived.date : '',
      col_map:            state.roster_col_map,
      seq_no:             state.current_seq_no,
    };

    const res = JSON.parse(await bridge.writeWorkResult(JSON.stringify(params)));
    if (!res.ok) {
      toast('명단 기록 실패: ' + res.error, 'err');
      return;
    }

    // 발송일 기록 (체크된 경우)
    if (sent.checked && sent.date) {
      const sentParams = {
        xlsx_path:   state.roster_log_path,
        school_name: state.selected_school,
        sent_date:   sent.date,
        col_map:     state.roster_col_map,
      };
      const sentRes = JSON.parse(await bridge.writeEmailSent(JSON.stringify(sentParams)));
      if (!sentRes.ok) {
        toast('작업 내용은 기록됐지만 발송일 기록 중 오류: ' + sentRes.error, 'warn');
      }
    }

    // 작업 이력 저장
    const counts = {};
    if (kindFlags.신입생) counts['신입생'] = (scanData?.items || []).find(i => i.kind === '신입생')?.row_count || 0;
    if (kindFlags.전입생) counts['전입생'] = (scanData?.items || []).find(i => i.kind === '전입생')?.row_count || 0;
    if (kindFlags.전출생) counts['전출생'] = (scanData?.items || []).find(i => i.kind === '전출생')?.row_count || 0;
    if (kindFlags.교직원) counts['교직원'] = (scanData?.items || []).find(i => i.kind === '교직원')?.row_count || 0;

    const entry = {
      last_date: state.work_date || _todayStr(),
      worker:    state.worker_name || '',
      counts,
    };
    await saveWorkHistoryEntry(entry);

    // UI 갱신
    state.pending_roster_log = false;
    setRosterBtns(false, true);

    // 이력 라벨 즉시 갱신
    const countStr = Object.entries(counts)
      .filter(([, v]) => v)
      .map(([k, v]) => `${k} ${v}명`)
      .join(' · ');
    const histText = `마지막 작업 · ${entry.last_date}` + (countStr ? `\n${countStr}` : '');
    const histEl = _el('school-history');
    if (histEl) { histEl.textContent = histText; histEl.style.display = ''; }

    // 버튼 텍스트 변경
    const btn = _el('btn-record-roster');
    if (btn) btn.textContent = '기록됨 ✓';

    toast(res.data.message || '명단 기록 완료', 'ok');
  }

  async function openRoster() {
    if (!state.roster_log_path) return;
    await bridge.openFile(state.roster_log_path);
  }

  // ──────────────────────────────────────────────
  // 학년도 아이디 규칙
  // ──────────────────────────────────────────────
  function _buildGradeRows() {
    const container = _el('grade-rows');
    if (!container) return;
    container.innerHTML = '';
    for (let g = 1; g <= 6; g++) {
      const row = document.createElement('div');
      row.className = 'grade-row-item';
      row.id        = `grade-row-${g}`;
      row.innerHTML = `
        <label>${g}학년</label>
        <input type="number" id="grade-year-${g}" min="0" max="2099" placeholder="-" value="" disabled>
      `;
      container.appendChild(row);
    }
  }

  function toggleGrade() {
    _gradeOpen = !_gradeOpen;
    _el('grade-map-body').classList.toggle('open', _gradeOpen);
    _el('btn-grade-toggle').textContent = _gradeOpen
      ? '학년도 아이디 규칙 숨기기 ▴'
      : '학년도 아이디 규칙 보기 ▾';
  }

  // state: "default" | "not_needed" | "no_roster" | "ok"
  function updateGradeMap(s, mapping) {
    for (let g = 1; g <= 6; g++) {
      const inp = _el(`grade-year-${g}`);
      if (!inp) continue;
      if (s === 'default' || s === 'not_needed') {
        inp.value    = '';
        inp.disabled = true;
      } else if (s === 'no_roster') {
        inp.value    = '';
        inp.disabled = false;
      } else if (s === 'ok') {
        const val    = mapping && mapping[g];
        inp.value    = val ? String(val) : '';
        inp.disabled = false;
      }
    }
  }

  // 학교명 기준 표시 학년 수 조정 (중/고 → 3, 나머지 → 6)
  function setGradeCount(schoolName) {
    const last = (schoolName || '').slice(-1);
    _maxGrade   = (last === '중' || last === '고') ? 3 : 6;
    for (let g = 1; g <= 6; g++) {
      const row = _el(`grade-row-${g}`);
      if (row) row.style.display = g <= _maxGrade ? '' : 'none';
      if (g > _maxGrade) {
        const inp = _el(`grade-year-${g}`);
        if (inp) inp.value = '';
      }
    }
  }

  function getGradeOverrides() {
    const result = {};
    for (let g = 1; g <= 6; g++) {
      const inp = _el(`grade-year-${g}`);
      if (!inp) continue;
      const v = parseInt(inp.value, 10);
      if (v > 0) result[g] = v;
    }
    return result;
  }

  function applyGrade() {
    const overrides = getGradeOverrides();
    if (!Object.keys(overrides).length) {
      toast('입력된 학년도 값이 없습니다.', 'info');
      return;
    }
    const lines = Object.entries(overrides)
      .sort((a, b) => a[0] - b[0])
      .map(([g, y]) => `${g}학년 → ${y}`);
    toast('학년도 아이디 규칙이 적용됩니다: ' + lines.join(', '), 'ok');
  }

  // ──────────────────────────────────────────────
  // 내부 헬퍼
  // ──────────────────────────────────────────────
  function _isExact(name) {
    return state.school_names.includes((name || '').trim());
  }

  function _setApplyBtn(enabled) {
    _el('btn-apply-school').disabled = !enabled;
  }

  function _setStatus(msg, cls) {
    const el = _el('school-status');
    if (!el) return;
    el.textContent = msg;
    el.className   = 'muted' + (cls ? ' ' + cls : '');
  }

  function _openDropdown()  { _el('school-dropdown').classList.add('open'); }
  function _closeDropdown() {
    _el('school-dropdown').classList.remove('open');
    _hlIndex = -1;
  }

  function _setDateInput(id, val) {
    const el = _el(id);
    if (el) el.value = val;
  }

  // 드롭다운 외부 클릭 시 닫기
  document.addEventListener('click', e => {
    if (!e.target.closest('#school-input') && !e.target.closest('#school-dropdown')) {
      _closeDropdown();
    }
  });

  // ──────────────────────────────────────────────
  // Public
  // ──────────────────────────────────────────────
  return {
    init, setWorkContext,
    onInput, onKeyDown, apply,
    updateSchoolInfo, reset, newSchool,
    getArrivedInfo, getSentInfo,
    setRosterBtns, recordRoster, openRoster,
    toggleGrade, updateGradeMap, setGradeCount,
    getGradeOverrides, applyGrade,
  };

})();
