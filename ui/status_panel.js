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

  let _searchTimer   = null;
  let _hlIndex       = -1;
  let _gradeOpen     = false;
  let _maxGrade      = 6;
  let _isKeyboardNav = false;

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
    if (_isKeyboardNav) return;
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
      el.addEventListener('mousedown', e => {
        e.preventDefault();
        _selectName(name);
      });
      dd.appendChild(el);
    });
  }

  function _selectName(name) {
    // 이전 검색 타이머 취소 — 타이머가 남아있으면 선택 후에 _applySearch가
    // 실행돼 버튼을 다시 disabled로 덮어쓰는 버그 방지
    clearTimeout(_searchTimer);
    _isKeyboardNav = true;
    _el('school-input').value = name;
    _isKeyboardNav = false;
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
      clearTimeout(_searchTimer);
      _openDropdown();
      _hlIndex = Math.min(_hlIndex + 1, items.length - 1);
      _highlight(items);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      clearTimeout(_searchTimer);
      _hlIndex = Math.max(_hlIndex - 1, 0);
      _highlight(items);
    } else if (e.key === 'Enter') {
      if (_hlIndex >= 0 && items[_hlIndex]) {
        // 드롭다운에서 항목 선택 후 엔터 → apply()와 동일하게 처리
        const name = items[_hlIndex].textContent;
        _isKeyboardNav = true;
        _el('school-input').value = name;
        _isKeyboardNav = false;
        _closeDropdown();
        _hlIndex = -1;
        // apply()를 직접 호출 → 검증·UI 갱신·스텝 전환 포함
        apply();
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
      _isKeyboardNav = true;
      inp.value = name;
      _isKeyboardNav = false;
      inp.setSelectionRange(name.length, name.length);
      _setApplyBtn(_isExact(name));
      _setStatus('적용 가능한 학교입니다.');
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
    else               { hist.textContent = '작업 이력 없음'; hist.style.display = ''; }

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
      toast('전체 명단 반영 실패: ' + res.error, 'err');
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

    // 작업 이력 업데이트 (전체 명단 반영 완료 상태만 갱신)
    const schoolYear = (state.work_date || _todayStr()).slice(0, 4);

    // 기존 이력 조회
    const histLoadRes = JSON.parse(await bridge.loadWorkHistory(schoolYear));
    const prevEntry = histLoadRes.ok
      ? (histLoadRes.data.history?.[state.selected_school] || {})
      : {};

    // 기존 실행 이력은 유지하고, 전체 명단 반영 상태만 추가
    const entry = {
      ...prevEntry,
      master_recorded: true,
      master_recorded_at: state.work_date || _todayStr(),
    };

    const _histRes = JSON.parse(
      await bridge.saveWorkHistory(
        schoolYear,
        state.selected_school,
        JSON.stringify(entry)
      )
    );

    if (!_histRes.ok) {
      console.error('[HISTORY] 업데이트 실패:', _histRes.error);
      toast('작업 이력 업데이트 오류: ' + (_histRes.error || ''), 'warn');
    } else {
      console.log('[HISTORY] 반영 완료:', schoolYear, state.selected_school);
    }

    // UI 갱신
    state.pending_roster_log = false;
    setRosterBtns(false, true);

    // 이력 라벨 즉시 갱신
    const SHORT = { '신입생': '신입', '전입생': '전입', '전출생': '전출', '교직원': '교직' };
    const countStr = Object.entries(counts)
      .filter(([, v]) => v)
      .map(([k, v]) => `${SHORT[k] ?? k} ${v}`)
      .join(' · ');

    let histText = `마지막 작업 · ${entry.last_date || (state.work_date || _todayStr())}`;
    if (entry.worker) histText += ` (${entry.worker})`;
    if (countStr) histText += `\n${countStr}`;
    histText += `\n전체 명단 반영 완료`;

    Panel.updateSchoolInfo({
      school_name: _el('current-school')?.textContent || state.selected_school,
      history_text: histText,
    });

    // 버튼 텍스트 변경
    const btn = _el('btn-record-roster');
    if (btn) btn.textContent = '반영 완료';

    toast(res.data.message || '전체 명단 반영 완료', 'ok');
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
        const val = mapping && (mapping[g] ?? mapping[String(g)]);
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
      return;
    }

    if (typeof Scan !== 'undefined' && Scan.applyManualGradeReady) {
      Scan.applyManualGradeReady(overrides);
    }
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
  function updateRosterMapBtn(rosterPath) {
    const btn  = _el('btn-open-roster-map');
    const hint = _el('roster-map-hint');
    if (!btn) return;
    if (rosterPath) {
      btn.disabled = false;
      btn._rosterPath = rosterPath;
      if (hint) hint.style.display = 'none';
    } else {
      btn.disabled = true;
      btn._rosterPath = null;
      if (hint) hint.style.display = '';
    }
  }

  async function openRosterMap() {
    const btn = _el('btn-open-roster-map');
    if (!btn || !btn._rosterPath) return;
    await bridge.openFile(btn._rosterPath);
  }

  return {
    init, setWorkContext,
    onInput, onKeyDown, apply,
    updateSchoolInfo, reset, newSchool,
    getArrivedInfo, getSentInfo,
    setRosterBtns, recordRoster, openRoster,
    toggleGrade, updateGradeMap, setGradeCount,
    getGradeOverrides, applyGrade,
    updateRosterMapBtn, openRosterMap,
  };

})();
