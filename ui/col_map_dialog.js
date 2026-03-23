/**
 * col_map_dialog.js — 명단 열 매핑 다이얼로그
 *
 * 의존: main.js (state, bridge, _el)
 *       setup.js (Setup.onColMapResult 콜백)
 *
 * 흐름:
 *   1. Setup.pickRoster() → 파일 선택 → ColMap.open(path, existingMap, onResult)
 *   2. bridge.readXlsxMeta() → 시트 목록 + 미리보기 로드
 *   3. 사용자가 역할 버튼 선택 → 헤더 열 클릭으로 매핑
 *   4. "설정 완료" → onResult(resultMap) 콜백 → setup.js에서 state 저장
 */

'use strict';

const ColMap = (() => {

  const ROLES = [
    { key: 'col_seq',       label: '자료실 순번',      required: false, color: '#F0F9FF' },
    { key: 'col_school',    label: '학교명',           required: true,  color: '#DBEAFE' },
    { key: 'col_domain',    label: '도메인(홈페이지)',  required: true,  color: '#E0E7FF' },
    { key: 'col_email_arr', label: '이메일 도착일자',  required: false, color: '#FEF9C3' },
    { key: 'col_email_snt', label: '완료 이메일 발송', required: false, color: '#FEF9C3' },
    { key: 'col_worker',    label: '작업자',           required: false, color: '#DCFCE7' },
    { key: 'col_freshmen',  label: '작업현황(신입생)', required: false, color: '#FFE4E6' },
    { key: 'col_transfer',  label: '작업현황(전입생)', required: false, color: '#FFE4E6' },
    { key: 'col_withdraw',  label: '작업현황(전출생)', required: false, color: '#FFE4E6' },
    { key: 'col_teacher',   label: '작업현황(교직원)', required: false, color: '#FFE4E6' },
  ];

  let _xlsxPath       = '';
  let _sheets         = [];
  let _headers        = [];
  let _rows           = [];
  let _colMap         = {};   // key → 0-based col index
  let _currentRoleIdx = 0;
  let _onResult       = null;

  // ──────────────────────────────────────────────
  // 열기
  // ──────────────────────────────────────────────
  async function open(xlsxPath, existingMap, onResult) {
    _xlsxPath       = xlsxPath;
    _colMap         = {};
    _onResult       = onResult;
    _currentRoleIdx = 0;

    // 기존 매핑 복원 (1-based → 0-based)
    ROLES.forEach(({ key }) => {
      const v = (existingMap || {})[key];
      if (v && v > 0) _colMap[key] = v - 1;
    });

    _renderRoleBtns();

    // 저장된 sheet / header_row 우선 적용
    const savedSheet  = (existingMap || {}).sheet      || '';
    const savedHeader = (existingMap || {}).header_row || 1;

    await _loadMeta(savedSheet, savedHeader);

    _el('cm-header-row').value = savedHeader;

    _selectRole(0);
    _refreshRoleBtns();
    _el('col-map-dialog').showModal();
  }

  // ──────────────────────────────────────────────
  // bridge.readXlsxMeta 호출
  // ──────────────────────────────────────────────
  async function _loadMeta(sheetName, headerRow) {
    const res = JSON.parse(
      await bridge.readXlsxMeta(_xlsxPath, sheetName || '', headerRow || 1)
    );
    if (!res.ok) { toast('파일을 읽을 수 없습니다: ' + res.error, 'err'); return; }

    _sheets  = res.data.sheets  || [];
    _headers = res.data.headers || [];
    _rows    = res.data.rows    || [];

    const sel = _el('cm-sheet');
    sel.innerHTML = _sheets.map(s =>
      `<option value="${_esc(s)}">${_esc(s)}</option>`
    ).join('');
    if (res.data.sheet) sel.value = res.data.sheet;

    _renderTable();
  }

  // ──────────────────────────────────────────────
  // 이벤트
  // ──────────────────────────────────────────────
  async function onSheetChange() {
    _colMap = {};
    await _loadMeta(_el('cm-sheet').value, parseInt(_el('cm-header-row').value, 10) || 1);
    _refreshRoleBtns();
    _selectRole(0);
  }

  async function onHeaderRowChange() {
    await _loadMeta(_el('cm-sheet').value, parseInt(_el('cm-header-row').value, 10) || 1);
  }

  // ──────────────────────────────────────────────
  // 역할 버튼
  // ──────────────────────────────────────────────
  function _renderRoleBtns() {
    const container = _el('cm-role-btns');
    container.textContent = '';
    ROLES.forEach((role, i) => {
      const btn = document.createElement('button');
      btn.id = `cm-rbtn-${i}`;
      btn.style.cssText = 'min-width:76px;height:52px;padding:4px 6px;background:#F8FAFC;border:1px solid #CBD5E1;border-radius:8px;font-size:11px;font-family:inherit;cursor:pointer;text-align:center;line-height:1.4;transition:background .1s,border .1s;';
      btn.textContent = role.label;
      const sub = document.createElement('span');
      sub.style.color = '#94A3B8';
      sub.textContent = '-';
      btn.appendChild(document.createElement('br'));
      btn.appendChild(sub);
      btn.addEventListener('click', () => selectRole(i));
      container.appendChild(btn);
    });
  }

  function _refreshRoleBtns() {
    ROLES.forEach((role, i) => {
      const btn = _el(`cm-rbtn-${i}`);
      if (!btn) return;
      const c      = _colMap[role.key];
      const active = i === _currentRoleIdx;
      const done   = c != null;

      const colLabel = done ? `${c + 1}열` : '-';
      btn.innerHTML = `${_esc(role.label)}<br><span style="font-weight:${done ? 700 : 400};color:${done ? '#1D4ED8' : '#94A3B8'}">${colLabel}</span>`;

      if (active) {
        btn.style.background = role.color || '#EFF6FF';
        btn.style.border     = '2px solid #3B82F6';
        btn.style.fontWeight = '700';
        btn.style.color      = '#0F172A';
      } else if (done) {
        btn.style.background = role.color || '#DBEAFE';
        btn.style.border     = '1px solid #CBD5E1';
        btn.style.fontWeight = '600';
        btn.style.color      = '#0F172A';
      } else {
        btn.style.background = '#F1F5F9';
        btn.style.border     = '1px solid #CBD5E1';
        btn.style.fontWeight = '400';
        btn.style.color      = '#94A3B8';
      }
    });
  }

  function selectRole(idx) {
    _currentRoleIdx = idx;
    _refreshRoleBtns();
    const role = ROLES[idx];
    const req  = role.required ? ' (필수)' : ' (선택)';
    _el('cm-guide').textContent = `▶  '${role.label}'${req} 열을 표 머리글에서 클릭해 주세요.`;
  }

  function _selectRole(idx) { selectRole(idx); }

  function skipRole() {
    const next = _currentRoleIdx + 1;
    if (next < ROLES.length) _selectRole(next);
    else _el('cm-guide').textContent = "✓  모든 항목 지정 완료. '설정 완료'를 눌러주세요.";
  }

  // ──────────────────────────────────────────────
  // 테이블 렌더
  // ──────────────────────────────────────────────
  function _renderTable() {
    const table = _el('cm-preview-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    const colColor = {};
    ROLES.forEach(role => {
      const c = _colMap[role.key];
      if (c != null) colColor[c] = role.color;
    });

    thead.innerHTML = '<tr>' + _headers.map((h, i) => {
      const bg      = colColor[i] || '#F8FAFC';
      const mapped  = colColor[i] != null;
      return `<th onclick="ColMap.onColClick(${i})" style="
        cursor:pointer;user-select:none;white-space:nowrap;
        padding:0;border:1px solid var(--border);
        background:${bg};position:sticky;top:0;min-width:60px;
      ">
        <div style="font-size:11px;font-weight:800;color:#334155;background:#E2E8F0;padding:2px 8px;border-bottom:1px solid var(--border);text-align:center;letter-spacing:.02em;">${i + 1}열</div>
        <div style="font-size:12px;font-weight:600;color:#0F172A;padding:5px 8px;text-align:center;">${_esc(h) || '<span style="color:#CBD5E1">-</span>'}</div>
      </th>`;
    }).join('') + '</tr>';

    tbody.innerHTML = _rows.map(row =>
      '<tr>' + _headers.map((_, i) => {
        const bg = colColor[i] ? `background:${colColor[i]}` : '';
        return `<td style="${bg};padding:5px 10px;border:1px solid var(--border);font-size:12px;white-space:nowrap">${_esc(row[i] ?? '')}</td>`;
      }).join('') + '</tr>'
    ).join('');
  }

  function onColClick(colIndex) {
    if (_currentRoleIdx >= ROLES.length) return;
    _colMap[ROLES[_currentRoleIdx].key] = colIndex;
    _renderTable();
    _refreshRoleBtns();
    const next = _currentRoleIdx + 1;
    if (next < ROLES.length) _selectRole(next);
    else _el('cm-guide').textContent = "✓  모든 항목 지정 완료. '설정 완료'를 눌러주세요.";
  }

  // ──────────────────────────────────────────────
  // 완료 / 취소
  // ──────────────────────────────────────────────
  function confirm() {
    const missing = ROLES.filter(r => r.required && _colMap[r.key] == null).map(r => r.label);
    if (missing.length) { toast('필수 항목 미지정: ' + missing.join(', '), 'warn'); return; }

    const headerRow = parseInt(_el('cm-header-row').value, 10) || 1;
    const result = { sheet: _el('cm-sheet').value, header_row: headerRow, data_start: headerRow + 1 };
    ROLES.forEach(({ key }) => { result[key] = (_colMap[key] != null) ? _colMap[key] + 1 : 0; });

    _el('col-map-dialog').close();
    if (_onResult) _onResult(result);
  }

  function stepHeaderRow(delta) {
    const el = _el('cm-header-row');
    const v = Math.min(30, Math.max(1, (parseInt(el.value, 10) || 1) + delta));
    el.value = v;
    _loadMeta(_el('cm-sheet').value, v);
  }

  function cancel() { _el('col-map-dialog').close(); }

  function _esc(str) {
    return String(str ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  return { open, selectRole, skipRole, onColClick, onSheetChange, onHeaderRowChange, stepHeaderRow, confirm, cancel };

})();
