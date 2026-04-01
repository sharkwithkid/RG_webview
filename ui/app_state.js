'use strict';

const state = {
  worker_name:        '',
  work_root:          '',
  work_date:          '',
  school_start_date:  '',
  roster_log_path:    '',
  roster_col_map:     {},
  arrived_date:       '',
  school_folders:     [],

  school_names:       [],
  selected_school:    '',
  selected_domain:    '',
  current_seq_no:     null,

  last_scan_logs:     [],
  last_run_logs:      [],
  last_diff_logs:     [],

  isInitializing:     true,
  isScanning:         false,
  isRunning:          false,
  isDiffScanning:     false,
  isDiffRunning:      false,
  isPreviewLoading:   false,

  currentPage:        'setup',
  currentTab:         'scan',
  currentMode:        'main',

  pending_roster_log:   false,
  school_kind_override: null,
};

const AppState = (() => {
  const LOG_KEY = {
    scan: 'last_scan_logs',
    run: 'last_run_logs',
    diff: 'last_diff_logs',
  };

  const BUSY_KEY = {
    scan:      'isScanning',
    run:       'isRunning',
    diff_scan: 'isDiffScanning',
    diff_run:  'isDiffRunning',
    preview:   'isPreviewLoading',
  };

  function setTaskLogs(kind, logs, fallbackMessage = '') {
    const key = LOG_KEY[kind];
    if (!key) return;
    const normalized = Array.isArray(logs) ? logs : [];
    state[key] = normalized.length
      ? normalized
      : (fallbackMessage ? [{ level: 'error', message: String(fallbackMessage) }] : []);
  }

  function clearTaskLogs() {
    state.last_scan_logs = [];
    state.last_run_logs = [];
    state.last_diff_logs = [];
  }

  function setBusy(kind, value) {
    const key = BUSY_KEY[kind];
    if (!key) return;
    state[key] = !!value;
  }

  function applySetup(params = {}) {
    state.work_root = params.work_root || '';
    state.roster_log_path = params.roster_log_path || '';
    state.worker_name = params.worker_name || '';
    state.school_start_date = params.school_start_date || '';
    state.work_date = params.work_date || '';
  }

  function setSelectedSchool(schoolName) {
    state.selected_school = schoolName || '';
    state.current_seq_no = null;
    state.pending_roster_log = false;
    state.school_kind_override = null;
  }

  function resetSelectionState() {
    state.selected_school = '';
    state.selected_domain = '';
    state.current_seq_no = null;
    state.pending_roster_log = false;
    state.school_kind_override = null;
  }

  function resetBusyFlags() {
    state.isScanning = false;
    state.isRunning = false;
    state.isDiffScanning = false;
    state.isDiffRunning = false;
    state.isPreviewLoading = false;
  }

  return {
    setTaskLogs,
    clearTaskLogs,
    setBusy,
    applySetup,
    setSelectedSchool,
    resetSelectionState,
    resetBusyFlags,
  };
})();
