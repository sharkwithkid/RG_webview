/**
 * date_picker.js — 커스텀 날짜 피커
 *
 * 사용법:
 *   DatePicker.init()  — DOMContentLoaded 후 자동 호출
 *
 * input[type="date"] 의 value를 읽고 쓰는 방식으로
 *   기존 코드(setup.js 등)와 완전히 호환됨.
 */

'use strict';

const DatePicker = (() => {

  const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];
  const MONTHS   = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];

  // fieldId → { input, trigger, popup, view, viewYear, viewMonth }
  const _pickers = {};
  let _openId = null;

  function init() {
    document.querySelectorAll('[data-dp]').forEach(trigger => {
      const fieldId = trigger.dataset.dp;
      const input   = document.getElementById(fieldId);
      const popup   = document.getElementById(`dp-popup-${fieldId}`);
      if (!input || !popup) return;

      const today = new Date();
      _pickers[fieldId] = {
        input,
        trigger,
        popup,
        view: 'days',       // 'days' | 'months' | 'years'
        viewYear:  today.getFullYear(),
        viewMonth: today.getMonth(),
      };

      trigger.addEventListener('click', e => {
        e.stopPropagation();
        _toggle(fieldId);
      });

      // 팝업 내부 클릭은 닫히지 않도록
      popup.addEventListener('click', e => e.stopPropagation());
    });

    document.addEventListener('click', _closeAll);
  }

  // ── 공개 API: 값 읽기/쓰기 ──────────────────────
  function getValue(fieldId) {
    return _pickers[fieldId]?.input.value || '';
  }

  function setValue(fieldId, dateStr) {
    const p = _pickers[fieldId];
    if (!p) return;
    p.input.value = dateStr;
    _updateTriggerText(fieldId, dateStr);
    p.input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // ── 내부 ─────────────────────────────────────────
  function _toggle(fieldId) {
    if (_openId && _openId !== fieldId) _close(_openId);
    if (_openId === fieldId) { _close(fieldId); return; }
    _open(fieldId);
  }

  function _open(fieldId) {
    const p = _pickers[fieldId];
    if (!p) return;

    // 현재 값으로 초기 뷰 설정
    const val = p.input.value;
    if (val) {
      const d = new Date(val + 'T00:00:00');
      p.viewYear  = d.getFullYear();
      p.viewMonth = d.getMonth();
    } else {
      const today = new Date();
      p.viewYear  = today.getFullYear();
      p.viewMonth = today.getMonth();
    }
    p.view = 'days';

    _renderDays(fieldId);

    // position: fixed로 팝업 위치 계산 (사이드바 overflow 무시)
    const rect = p.trigger.getBoundingClientRect();
    const popup = p.popup;
    popup.style.position = 'fixed';
    popup.style.left = rect.left + 'px';
    popup.style.minWidth = '260px';

    const spaceBelow = window.innerHeight - rect.bottom;
    if (spaceBelow < 300) {
      popup.style.bottom = (window.innerHeight - rect.top + 4) + 'px';
      popup.style.top = 'auto';
    } else {
      popup.style.top = (rect.bottom + 4) + 'px';
      popup.style.bottom = 'auto';
    }

    popup.classList.add('open');
    p.trigger.classList.add('open');
    _openId = fieldId;
  }

  function _close(fieldId) {
    const p = _pickers[fieldId];
    if (!p) return;
    p.popup.classList.remove('open');
    p.trigger.classList.remove('open');
    if (_openId === fieldId) _openId = null;
  }

  function _closeAll() {
    if (_openId) _close(_openId);
  }

  // ── 렌더링 ────────────────────────────────────────
  function _renderDays(fieldId) {
    const p   = _pickers[fieldId];
    const sel = p.input.value ? new Date(p.input.value + 'T00:00:00') : null;
    const today = new Date();
    const y = p.viewYear, m = p.viewMonth;

    const firstDay = new Date(y, m, 1).getDay();
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const daysInPrev  = new Date(y, m, 0).getDate();

    let html = `
      <div class="dp-nav">
        <button class="dp-nav-btn" onclick="DatePicker._prevMonth('${fieldId}')">&#8249;</button>
        <span class="dp-nav-label" onclick="DatePicker._switchView('${fieldId}','months')">${y}년 ${MONTHS[m]}</span>
        <button class="dp-nav-btn" onclick="DatePicker._nextMonth('${fieldId}')">&#8250;</button>
      </div>
      <div class="dp-weekdays">${WEEKDAYS.map(w => `<div class="dp-weekday">${w}</div>`).join('')}</div>
      <div class="dp-days">`;

    // 이전 달
    for (let i = firstDay - 1; i >= 0; i--) {
      html += `<button class="dp-day other-month" onclick="DatePicker._prevMonth('${fieldId}');DatePicker._selectDay('${fieldId}',${y},${m-1},${daysInPrev-i})">${daysInPrev - i}</button>`;
    }
    // 이번 달
    for (let d = 1; d <= daysInMonth; d++) {
      const isToday    = y === today.getFullYear() && m === today.getMonth() && d === today.getDate();
      const isSelected = sel && y === sel.getFullYear() && m === sel.getMonth() && d === sel.getDate();
      const cls = ['dp-day', isToday ? 'today' : '', isSelected ? 'selected' : ''].filter(Boolean).join(' ');
      html += `<button class="${cls}" onclick="DatePicker._selectDay('${fieldId}',${y},${m},${d})">${d}</button>`;
    }
    // 다음 달
    const remaining = 42 - firstDay - daysInMonth;
    for (let d = 1; d <= remaining; d++) {
      html += `<button class="dp-day other-month" onclick="DatePicker._nextMonth('${fieldId}');DatePicker._selectDay('${fieldId}',${y},${m+1},${d})">${d}</button>`;
    }

    html += `</div>`;
    p.popup.innerHTML = html;
  }

  function _renderMonths(fieldId) {
    const p = _pickers[fieldId];
    const sel = p.input.value ? new Date(p.input.value + 'T00:00:00') : null;

    let html = `
      <div class="dp-nav">
        <button class="dp-nav-btn" onclick="DatePicker._prevYear('${fieldId}')">&#8249;</button>
        <span class="dp-nav-label" onclick="DatePicker._switchView('${fieldId}','years')">${p.viewYear}년</span>
        <button class="dp-nav-btn" onclick="DatePicker._nextYear('${fieldId}')">&#8250;</button>
      </div>
      <div class="dp-month-grid">`;

    MONTHS.forEach((name, i) => {
      const isSel = sel && p.viewYear === sel.getFullYear() && i === sel.getMonth();
      html += `<button class="dp-month-btn${isSel ? ' selected' : ''}" onclick="DatePicker._selectMonth('${fieldId}',${i})">${name}</button>`;
    });
    html += `</div>`;
    p.popup.innerHTML = html;
  }

  function _renderYears(fieldId) {
    const p   = _pickers[fieldId];
    const sel = p.input.value ? new Date(p.input.value + 'T00:00:00') : null;
    const base = Math.floor(p.viewYear / 12) * 12;

    let html = `
      <div class="dp-nav">
        <button class="dp-nav-btn" onclick="DatePicker._prevYearPage('${fieldId}')">&#8249;</button>
        <span class="dp-nav-label">${base}–${base+11}</span>
        <button class="dp-nav-btn" onclick="DatePicker._nextYearPage('${fieldId}')">&#8250;</button>
      </div>
      <div class="dp-year-grid">`;

    for (let y = base; y < base + 12; y++) {
      const isSel = sel && y === sel.getFullYear();
      html += `<button class="dp-year-btn${isSel ? ' selected' : ''}" onclick="DatePicker._selectYear('${fieldId}',${y})">${y}</button>`;
    }
    html += `</div>`;
    p.popup.innerHTML = html;
  }

  // ── 이벤트 핸들러 (전역 노출) ─────────────────────
  function _switchView(fieldId, view) {
    _pickers[fieldId].view = view;
    if (view === 'months') _renderMonths(fieldId);
    else if (view === 'years') _renderYears(fieldId);
    else _renderDays(fieldId);
  }

  function _prevMonth(fieldId) {
    const p = _pickers[fieldId];
    if (--p.viewMonth < 0) { p.viewMonth = 11; p.viewYear--; }
    _renderDays(fieldId);
  }

  function _nextMonth(fieldId) {
    const p = _pickers[fieldId];
    if (++p.viewMonth > 11) { p.viewMonth = 0; p.viewYear++; }
    _renderDays(fieldId);
  }

  function _prevYear(fieldId) { _pickers[fieldId].viewYear--; _renderMonths(fieldId); }
  function _nextYear(fieldId) { _pickers[fieldId].viewYear++; _renderMonths(fieldId); }
  function _prevYearPage(fieldId) { _pickers[fieldId].viewYear -= 12; _renderYears(fieldId); }
  function _nextYearPage(fieldId) { _pickers[fieldId].viewYear += 12; _renderYears(fieldId); }

  function _selectDay(fieldId, y, m, d) {
    const real_m = ((m % 12) + 12) % 12;
    const real_y = y + Math.floor(m / 12);
    const dateStr = `${real_y}-${String(real_m + 1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    setValue(fieldId, dateStr);
    _close(fieldId);
  }

  function _selectMonth(fieldId, m) {
    _pickers[fieldId].viewMonth = m;
    _switchView(fieldId, 'days');
  }

  function _selectYear(fieldId, y) {
    _pickers[fieldId].viewYear = y;
    _switchView(fieldId, 'months');
  }

  function _updateTriggerText(fieldId, dateStr) {
    const p = _pickers[fieldId];
    if (!p) return;
    const span = p.trigger.querySelector('.dp-text');
    if (!span) return;
    if (dateStr) {
      const d = new Date(dateStr + 'T00:00:00');
      span.textContent = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      span.classList.remove('dp-placeholder');
    } else {
      span.textContent = '날짜 선택';
      span.classList.add('dp-placeholder');
    }
  }

  // setup.js의 _setVal이 input.value를 직접 바꾸므로
  // MutationObserver 대신 input 이벤트로 트리거 텍스트 동기화
  function _syncAll() {
    Object.keys(_pickers).forEach(fieldId => {
      const val = _pickers[fieldId].input.value;
      if (val) _updateTriggerText(fieldId, val);
    });
  }

  return {
    init, getValue, setValue, _syncAll,
    _switchView, _prevMonth, _nextMonth,
    _prevYear, _nextYear, _prevYearPage, _nextYearPage,
    _selectDay, _selectMonth, _selectYear,
  };
})();
