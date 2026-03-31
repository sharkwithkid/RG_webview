'use strict';

const UICommon = (() => {
  function getStatusMessages(status, levels = null) {
    const msgs = Array.isArray(status?.messages) ? status.messages : [];
    return msgs
      .filter(m => !levels || levels.includes(m.level))
      .map(m => String(m?.text || '').trim())
      .filter(Boolean);
  }

  function getStatusDetailMessages(status) {
    const details = Array.isArray(status?.detail_messages) ? status.detail_messages : [];
    return details.map(v => String(v || '').trim()).filter(Boolean);
  }

  function getLogMessages(logs, levels = ['warn', 'error']) {
    return (Array.isArray(logs) ? logs : [])
      .filter(l => levels.includes(String(l?.level || '').toLowerCase()))
      .map(l => String(l?.message || '').trim())
      .filter(Boolean);
  }

  function collectMessages({ status = null, events = [], logs = [], prefer = ['error', 'hold', 'warn'], allowLogs = false } = {}) {
    const messages = [];
    const seen = new Set();
    prefer.forEach(level => {
      getStatusMessages(status, [level]).forEach(msg => {
        if (!seen.has(msg)) {
          seen.add(msg);
          messages.push(msg);
        }
      });
    });
    if (allowLogs && !messages.length) {
      getLogMessages(logs).forEach(msg => {
        if (!seen.has(msg)) {
          seen.add(msg);
          messages.push(msg);
        }
      });
    }
    if (!messages.length && Array.isArray(events)) {
      events.map(e => String(e?.message || '').trim()).filter(Boolean).forEach(msg => {
        if (!seen.has(msg)) {
          seen.add(msg);
          messages.push(msg);
        }
      });
    }
    return messages;
  }

  function primaryMessage({ status = null, events = [], logs = [], prefer = ['error', 'hold', 'warn'], allowLogs = false } = {}) {
    return collectMessages({ status, events, logs, prefer, allowLogs })[0] || '';
  }

  function renderStatusCard(elOrId, messages = [], mode = 'warn', status = null) {
    const el = typeof elOrId === 'string' ? _el(elOrId) : elOrId;
    if (!el) return;
    StatusUI.renderWarnCard(el, messages, mode, status);
  }

  function hideStatusCard(elOrId) {
    const el = typeof elOrId === 'string' ? _el(elOrId) : elOrId;
    if (!el) return;
    el.style.display = 'none';
    el.innerHTML = '';
    el.classList.remove('error');
  }

  function renderWarnCard(elOrId, status, mode = 'warn', fallbackMessages = []) {
    const el = typeof elOrId === 'string' ? _el(elOrId) : elOrId;
    if (!el) return;
    const messages = collectMessages({ status, logs: [], events: [], prefer: mode === 'error' ? ['error'] : ['warn', 'hold'] });
    const source = messages.length ? messages : (Array.isArray(fallbackMessages) ? fallbackMessages : []);
    renderStatusCard(el, source, mode, status);
  }

  function renderOutputFiles(elOrId, files, onOpen) {
    const el = typeof elOrId === 'string' ? _el(elOrId) : elOrId;
    if (!el) return;
    el.textContent = '';
    const safeFiles = Array.isArray(files) ? files : [];
    if (!safeFiles.length) {
      const empty = document.createElement('span');
      empty.className = 'muted';
      empty.textContent = '생성된 파일 없음';
      el.appendChild(empty);
      return;
    }
    safeFiles.forEach(file => {
      const row = document.createElement('div');
      row.className = 'output-file-item';
      const link = document.createElement('span');
      link.className = 'output-file-name';
      link.textContent = file.name;
      link.addEventListener('click', () => onOpen?.(file));
      row.appendChild(link);
      el.appendChild(row);
    });
  }

  function subtractMessage(messages, text) {
    return (Array.isArray(messages) ? messages : []).filter(msg => String(msg || '').trim() !== String(text || '').trim());
  }

  return {
    getStatusMessages,
    getStatusDetailMessages,
    getLogMessages,
    collectMessages,
    primaryMessage,
    renderStatusCard,
    hideStatusCard,
    renderWarnCard,
    renderOutputFiles,
    subtractMessage,
  };
})();
