/**
 * notice_tab.js — 안내문 탭 로직
 *
 * 의존: main.js (state, bridge, _el)
 * HTML ID: notice-list, notice-text
 *
 * 템플릿 플레이스홀더:
 *   {school_name}, {year}, {prev_year}, {month}, {day}, {domain}
 */

'use strict';

const Notice = (() => {

  let _templates    = {};   // { key: rawText }
  let _rendered     = {};   // { key: replacedText }
  let _currentKey   = null;
  let _lastCtx      = null;

  // ──────────────────────────────────────────────
  // 템플릿 로드 (App.onSchoolSelected에서 호출)
  // ──────────────────────────────────────────────
  function loadTemplates(rawTemplates, ctx) {
    _templates  = rawTemplates || {};
    _rendered   = {};
    _currentKey = null;
    _lastCtx    = ctx || {};
    _applyReplace(_lastCtx);
    _renderList();
    _selectFirst();
    // 캡션 업데이트
    const captionEl = document.querySelector('#tab-notice .muted');
    if (captionEl) captionEl.textContent = '학교명·학년도·개학일·홈페이지 주소가 자동으로 반영된 안내문입니다.';
  }

  // ──────────────────────────────────────────────
  // 플레이스홀더 치환
  // ──────────────────────────────────────────────
  function _applyReplace(ctx) {
    // ctx: { school_name, domain, school_start_date } — state 직접 참조 안 함
    const c      = ctx || {};
    const school = c.school_name       || '';
    const domain = c.domain            || '';
    const date   = c.school_start_date || '';
    const year   = date.slice(0, 4);
    const prevY  = year ? String(parseInt(year, 10) - 1) : '';
    const month  = date.slice(5, 7)?.replace(/^0/, '') || '';
    const day    = date.slice(8, 10)?.replace(/^0/, '') || '';

    // 정렬: 신규등록메일 > 신규등록문자 > 반이동메일(신입생교직원포함) > 반이동메일 > 반이동문자 > 교직원 > 2-6메일 > 2-6문자
    const ORDER = [
      '신규등록 - 메일',
      '신규등록 - 문자',
      '반이동 - 메일 (신입생',   // "반이동 - 메일 (신입생, 교직원 등록 & 반이동)" 포함
      '반이동 - 메일',
      '반이동 - 문자',
      '교직원',
      '2-6학년 명단 보내 온 경우 - 메일',
      '2-6학년',
    ];
    const sortKey = k => {
      const idx = ORDER.findIndex(kw => k.includes(kw));
      return idx >= 0 ? idx : ORDER.length;
    };

    const sorted = Object.keys(_templates).sort((a, b) => {
      const d = sortKey(a) - sortKey(b);
      return d !== 0 ? d : a.localeCompare(b, 'ko');
    });

    _rendered = {};
    sorted.forEach(key => {
      _rendered[key] = (_templates[key] || '')
        .replace(/\{school_name\}/g, school)
        .replace(/\{year\}/g,        year)
        .replace(/\{prev_year\}/g,   prevY)
        .replace(/\{month\}/g,       month)
        .replace(/\{day\}/g,         day)
        .replace(/\{domain\}/g,      domain);
    });
  }

  // ──────────────────────────────────────────────
  // 목록 렌더링
  // ──────────────────────────────────────────────
  function _renderList() {
    const el = _el('notice-list');
    if (!el) return;

    el.textContent = '';  // innerHTML 대신 자식 제거

    const keys = Object.keys(_rendered);
    if (!keys.length) {
      const empty = document.createElement('div');
      empty.className   = 'notice-list-item muted';
      empty.textContent = '안내문 없음';
      el.appendChild(empty);
      return;
    }

    keys.forEach(key => {
      const item = document.createElement('div');
      item.className    = 'notice-list-item';
      item.dataset.key  = key;
      item.textContent  = key;
      item.addEventListener('click', () => selectKey(key));
      el.appendChild(item);
    });
  }

  function _selectFirst() {
    const keys = Object.keys(_rendered);
    if (keys.length) selectKey(keys[0]);
  }

  // ──────────────────────────────────────────────
  // 항목 선택
  // ──────────────────────────────────────────────
  function selectKey(key) {
    _currentKey = key;
    document.querySelectorAll('.notice-list-item').forEach(el => {
      el.classList.toggle('active', el.dataset.key === key);
    });
    const text = _rendered[key] || '';
    const el = _el('notice-text');
    if (el) el.value = text;
  }

  // ──────────────────────────────────────────────
  // 복사
  // ──────────────────────────────────────────────
  async function copy() {
    const el = _el('notice-text');
    const text = el?.value || '';
    if (!text) { toast('복사할 내용이 없습니다.', 'info'); return; }
    await bridge.copyToClipboard(text);
    // 짧은 피드백 (버튼 텍스트 변경)
    const btn = document.querySelector('[onclick="Notice.copy()"]');
    if (btn) {
      const orig = btn.textContent;
      btn.textContent = '복사됨 ✓';
      setTimeout(() => { btn.textContent = orig; }, 1200);
    }
  }

  // ──────────────────────────────────────────────
  // 초기화 (원본 치환값으로 복원)
  // ──────────────────────────────────────────────
  function reset() {
    if (!_currentKey) return;
    _applyReplace(_lastCtx || {});
    selectKey(_currentKey);
  }

  // ──────────────────────────────────────────────
  // 학교/도메인 변경 시 전체 재치환 (App.onSchoolSelected 후 호출 가능)
  // ──────────────────────────────────────────────
  function refresh(ctx) {
    if (!Object.keys(_templates).length) return;
    _lastCtx = ctx || _lastCtx || {};
    _applyReplace(_lastCtx);
    _renderList();
    if (_currentKey) selectKey(_currentKey);
    else _selectFirst();
  }

  // ──────────────────────────────────────────────
  // 학교 변경 시 전체 초기화
  // ──────────────────────────────────────────────
  function clear() {
    _templates  = {};
    _rendered   = {};
    _currentKey = null;
    _lastCtx    = null;
    const listEl = _el('notice-list');
    if (listEl) listEl.innerHTML = '';
    const textEl = _el('notice-text');
    if (textEl) textEl.value =
      '학교를 선택하고 실행을 완료하면\n안내문이 자동으로 채워집니다.\n\n' +
      '학교명 · 학년도 · 개학일 · 홈페이지 주소가 자동 반영됩니다.';
  }

  function _esc(str) {
    return String(str ?? '')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  return { loadTemplates, selectKey, copy, reset, refresh, clear };

})();
