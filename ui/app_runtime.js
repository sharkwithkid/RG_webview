/**
 * app_runtime.js — Bridge signal 처리 + 상위 워크플로우 orchestration
 *
 * 원칙:
 *  - main.js에서 비대한 async workflow / signal wiring 분리
 *  - 상태 전이는 AppState를 통해 공통화
 *  - 화면 이동/렌더링은 기존 App/Panel/Scan/Run/Diff/Notice에 위임
 */
'use strict';

const BridgeRuntime = (() => {
  function _taskFailedMessage(kind) {
    return {
      scan: '예기치 못한 오류가 발생했습니다. 스캔 로그를 확인해 주세요.',
      run: '예기치 못한 오류가 발생했습니다. 실행 로그를 확인해 주세요.',
      diff: '예기치 못한 오류가 발생했습니다. 로그를 확인해 주세요.',
    }[kind] || '예기치 못한 오류가 발생했습니다.';
  }

  function _setTaskCompleted(kind, payload, handler) {
    AppState.setBusy(kind, false);
    AppState.setTaskLogs(kind, payload?.data?.logs || []);
    if (payload?.ok) handler.success(payload.data);
    else handler.fail(_taskFailedMessage(kind), payload?.error || '');
  }

  function connectSignals(activeBridge) {
    activeBridge.scanFinished.connect(payload => {
      const p = JSON.parse(payload);
      _setTaskCompleted('scan', p, {
        success: data => Scan.onFinished(data),
        fail: message => Scan.onFailed(message),
      });
    });

    activeBridge.scanFailed.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('scan', false);
      AppState.setTaskLogs('scan', [], p.error || '');
      Scan.onFailed(_taskFailedMessage('scan'));
    });

    activeBridge.runFinished.connect(payload => {
      const p = JSON.parse(payload);
      _setTaskCompleted('run', p, {
        success: data => Run.onFinished(data),
        fail: message => Run.onFailed(message),
      });
    });

    activeBridge.runFailed.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('run', false);
      AppState.setTaskLogs('run', [], p.error || '');
      Run.onFailed(_taskFailedMessage('run'));
    });

    activeBridge.diffScanFinished.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('diff_scan', false);
      if (p.ok) Diff.onScanFinished(p.data);
      else Diff.onScanFailed(p.error || '명단 비교 스캔 실패');
    });

    activeBridge.diffScanFailed.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('diff_scan', false);
      Diff.onScanFailed(p.error || '예기치 못한 오류');
    });

    activeBridge.diffRunFinished.connect(payload => {
      const p = JSON.parse(payload);
      _setTaskCompleted('diff', p, {
        success: data => Diff.onFinished(data),
        fail: message => Diff.onFailed(message),
      });
    });

    activeBridge.diffRunFailed.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('diff_run', false);
      AppState.setTaskLogs('diff', [], p.error || '');
      Diff.onFailed(_taskFailedMessage('diff'));
    });

    activeBridge.previewLoaded.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('preview', false);
      if (p.ok) {
        if (p.kind === 'run_output') Run.onPreviewLoaded(p);
        else if (p.kind === 'compare' || p.kind === '재학생') Diff.onPreviewLoaded?.(p);
        else Scan.onPreviewLoaded(p);
        return;
      }
      if (p.kind === 'run_output') _el('run-preview-info').textContent = p.error || '미리보기 실패';
      else if (p.kind === 'compare' || p.kind === '재학생') Diff.onPreviewFailed?.(p.kind, p.error || '미리보기 실패');
      else Scan.onPreviewFailed(p.kind, p.error || '미리보기 실패');
    });

    activeBridge.previewFailed.connect(payload => {
      const p = JSON.parse(payload);
      AppState.setBusy('preview', false);
      if (p.kind === 'run_output') Run.onPreviewFailed?.(p.kind || '', p.error || '예기치 못한 오류');
      else if (p.kind === 'compare' || p.kind === '재학생') Diff.onPreviewFailed?.(p.kind || '', p.error || '예기치 못한 오류');
      else Scan.onPreviewFailed(p.kind || '', p.error || '예기치 못한 오류');
    });
  }

  return { connectSignals };
})();

const Workflow = (() => {
  async function onSetupComplete(params, activeBridge) {
    AppState.applySetup(params);
    await activeBridge.saveAppConfig(JSON.stringify(_readFullConfig()));

    _el('header-work-date').textContent = `작업일 · ${state.work_date}`;

    const inspRes = JSON.parse(await activeBridge.inspectWorkRoot(state.work_root));
    state.school_folders = inspRes.ok ? (inspRes.data.school_folders || []) : [];

    const namesRes = JSON.parse(
      await activeBridge.loadSchoolNames(
        state.roster_log_path,
        JSON.stringify(state.roster_col_map || {})
      )
    );
    state.school_names = namesRes.ok ? (namesRes.data.school_names || []) : [];
    Panel.init(state.school_names);
    Panel.setWorkContext({ work_date: state.work_date, arrived_date: state.arrived_date });

    App.setStepState(0, 'done');
    App.setStepState(1, 'active');
    for (let i = 2; i <= 4; i++) App.setStepState(i, 'idle');

    _showPage('main');
    state.currentTab = 'scan';
    document.querySelectorAll('.tab-panel').forEach(el => el.classList.remove('active'));
    _el('tab-scan')?.classList.add('active');
    _highlightStep(1);
    state.currentMode = 'main';
    _el('btn-mode-main').classList.add('active');
    _el('btn-mode-diff').classList.remove('active');
    _setSchoolInputHighlight(true);
  }

  async function onSchoolSelected(schoolName, activeBridge) {
    const prevSchool = state.selected_school || '';
    const schoolChanged = !!prevSchool && prevSchool !== schoolName;

    if (typeof Scan !== 'undefined' && Scan.reset) Scan.reset();
    if (typeof Run !== 'undefined' && Run.reset) Run.reset();
    if (typeof Diff !== 'undefined' && Diff.reset) Diff.reset();
    if (schoolChanged && typeof Panel !== 'undefined' && Panel.resetSchoolContext) {
      Panel.resetSchoolContext();
    }

    AppState.setSelectedSchool(schoolName);

    if (state.work_root) {
      try {
        const inspRes = JSON.parse(await activeBridge.inspectWorkRoot(state.work_root));
        if (inspRes.ok && inspRes.data?.ok) {
          state.school_folders = inspRes.data.school_folders || [];
        }
      } catch (e) {
        console.warn('inspectWorkRoot refresh failed:', e);
      }
    }

    const domRes = JSON.parse(
      await activeBridge.getSchoolDomain(
        state.roster_log_path,
        schoolName,
        JSON.stringify(state.roster_col_map || {})
      )
    );
    state.selected_domain = domRes.ok ? (domRes.data.domain || '') : '';

    const tplRes = JSON.parse(await activeBridge.loadNoticeTemplates(state.work_root));
    if (tplRes.ok) Notice.loadTemplates(tplRes.data.templates || {}, _noticeCtx());

    const schoolYear = (state.work_date || _todayStr()).slice(0, 4);
    const histRes = JSON.parse(await activeBridge.loadWorkHistory(schoolYear));
    const histEntry = histRes.ok ? (histRes.data.history?.[schoolName] || null) : null;

    let history_text = null;
    if (histEntry) {
      const SHORT = { '신입생': '신입', '전입생': '전입', '전출생': '전출', '교직원': '교직' };
      const countStr = Object.entries(histEntry.counts || {})
        .filter(([, v]) => v)
        .map(([k, v]) => `${SHORT[k] ?? k} ${v}`)
        .join(' · ');
      history_text = `마지막 작업 · ${histEntry.last_date || '-'}`;
      if (histEntry.worker) history_text += ` (${histEntry.worker})`;
      if (countStr) history_text += `\n${countStr}`;
    }

    const normalize = s => (s && s.normalize ? s.normalize('NFC') : (s || ''));
    const compact = s => normalize(s).replace(/[\s._-]+/g, '');
    const schoolNameCompact = compact(schoolName);
    const folderName = state.school_folders.find(f => {
      const nf = normalize(f);
      return nf.includes(normalize(schoolName)) || compact(nf).includes(schoolNameCompact);
    }) || schoolName;

    Panel.updateSchoolInfo({ school_name: folderName, history_text });
    Panel.setGradeCount(schoolName);

    App.setStepState(1, 'done');
    App.setStepState(2, 'active');
    for (let i = 3; i <= 4; i++) App.setStepState(i, 'idle');

    _el('btn-scan').disabled = false;
    _el('btn-run').disabled = true;
    Panel.setRosterBtns(false, !!state.roster_log_path);
    _setSchoolInputHighlight(false);

    if (state.currentMode === 'diff') {
      App.goTab('diff');
      Diff.reset();
      _el('btn-scan-diff').disabled = false;
      _el('btn-run-diff').disabled = true;
    } else {
      _el('btn-scan-diff').disabled = true;
      _el('btn-run-diff').disabled = true;
      App.goTab('scan');
      Scan.reset();
    }
  }

  function resetSharedState() {
    AppState.resetSelectionState();
    AppState.resetBusyFlags();
    AppState.clearTaskLogs();

    Panel.reset();
    Scan.reset();
    Run.reset();
    if (typeof Diff !== 'undefined' && Diff.reset) Diff.reset();
    Notice.clear();
    App.setFloatingNext(false, null);

    const scanBadge = _el('scan-status-badge');
    if (scanBadge) { scanBadge.className = 'status-badge badge-idle'; scanBadge.textContent = '대기'; }
    const runBadge = _el('run-status-badge');
    if (runBadge) { runBadge.className = 'status-badge badge-idle'; runBadge.textContent = '대기'; }
    const btnGotoRun = _el('btn-goto-run');
    if (btnGotoRun) btnGotoRun.style.display = 'none';
    const btnGotoNotice = _el('btn-goto-notice');
    if (btnGotoNotice) btnGotoNotice.style.display = 'none';

    _el('btn-scan').disabled = true;
    _el('btn-run').disabled = true;
    const btnScanDiff = _el('btn-scan-diff');
    if (btnScanDiff) btnScanDiff.disabled = true;
    _el('btn-run-diff').disabled = true;
  }

  return { onSetupComplete, onSchoolSelected, resetSharedState };
})();
