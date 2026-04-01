from __future__ import annotations

from typing import Any

from core.common import derive_grade_year_map
from core.events import CoreEvent, RowMark


def parse_log_entry(raw: str) -> dict:
    import re as _re
    s = str(raw)
    s = _re.sub(r'^\[\d{2}:\d{2}:\d{2}\]\s*', '', s)
    if s.startswith('[ERROR]'):
        return {'level': 'error', 'message': s[7:].strip()}
    if s.startswith('[WARN]'):
        return {'level': 'warn', 'message': s[6:].strip()}
    if s.startswith('[DEBUG]'):
        return {'level': 'debug', 'message': s[7:].strip()}
    if s.startswith('[INFO]'):
        return {'level': 'info', 'message': s[6:].strip()}
    if s.startswith('[DONE]'):
        return {'level': 'info', 'message': s[6:].strip()}
    if s.startswith('[TIMER]'):
        return {'level': 'info', 'message': s.strip()}
    return {'level': 'info', 'message': s.strip()}


def logs_from_result(result: Any) -> list[dict]:
    return [parse_log_entry(l) for l in (getattr(result, 'logs', None) or [])]


def _badge_for_level(level: str) -> dict:
    mapping = {
        'error': {'type': 'err', 'text': '오류'},
        'hold': {'type': 'hold', 'text': '보류'},
        'warn': {'type': 'warn', 'text': '경고'},
        'ok': {'type': 'ok', 'text': '완료'},
    }
    return mapping.get(level, mapping['ok'])


def _summary_for_level(level: str, count: int, ok_text: str = '완료') -> str:
    if level == 'error':
        return f'오류 {count}건이 있습니다.'
    if level == 'hold':
        return f'보류 {count}건이 있습니다.'
    if level == 'warn':
        return f'경고 {count}건이 있습니다.'
    return ok_text


def build_status(level: str, messages: list[dict] | None = None, *, summary_text: str | None = None,
                 detail_messages: list[str] | None = None, action_text: str = '', row_marks: dict | None = None) -> dict:
    safe_messages = list(messages or [])
    # detail_messages 미지정 시 messages의 text 필드로 일관되게 생성
    # (messages 필드명이 'text'로 통일되어 있으므로 'message' fallback 불필요)
    safe_details = list(detail_messages if detail_messages is not None else [m.get('text', '') for m in safe_messages if m.get('text')])
    return {
        'level': level,
        'badge': _badge_for_level(level),
        'messages': safe_messages,
        'summary_text': summary_text or _summary_for_level(level, len(safe_messages)),
        'detail_messages': safe_details,
        'action_text': action_text,
        'row_marks': row_marks or {'warn_rows': [], 'error_rows': [], 'issue_rows': []},
    }


def status_from_events(events: list[CoreEvent]) -> dict:
    errs = [e for e in events if e.level == 'error']
    holds = [e for e in events if e.level == 'hold']
    warns = [e for e in events if e.level == 'warn']
    if errs:
        level, selected = 'error', errs
    elif holds:
        level, selected = 'hold', holds
    elif warns:
        level, selected = 'warn', warns
    else:
        level, selected = 'ok', []
    messages = [{'level': e.level, 'text': e.message, 'code': e.code} for e in (errs + holds + warns)]
    return build_status(
        level,
        messages,
        summary_text=_summary_for_level(level, len(selected)),
        detail_messages=[e.message for e in (errs + holds + warns)],
    )


def serialize_event(e: CoreEvent) -> dict:
    return {
        'code': e.code,
        'level': e.level,
        'message': e.message,
        'detail': e.detail,
        'file_key': e.file_key,
        'row': e.row,
        'field_name': e.field_name,
        'blocking': e.blocking,
    }


def serialize_row_mark(m: RowMark) -> dict:
    return {'file_key': m.file_key, 'row': m.row, 'level': m.level, 'code': m.code}


def meta_get(meta, key, default=None):
    if meta is None:
        return default
    if isinstance(meta, dict):
        return meta.get(key, default)
    return getattr(meta, key, default)


def as_abs_issue_rows(issue_rows, data_start_row):
    out = []
    base = int(data_start_row or 1)
    for r in list(issue_rows or []):
        try:
            r_i = int(r)
        except Exception:
            continue
        out.append(r_i if r_i >= base else base + r_i)
    seen = set()
    uniq = []
    for r in out:
        if r not in seen:
            seen.add(r)
            uniq.append(r)
    return uniq


def _row_mark_summary(row_marks: list[RowMark]) -> dict:
    warn_rows = [m.row for m in row_marks if m.level in ('warn', 'dup', 'hold')]
    error_rows = [m.row for m in row_marks if m.level == 'error']
    issue_rows = warn_rows + error_rows
    return {'warn_rows': warn_rows, 'error_rows': error_rows, 'issue_rows': issue_rows}


def _fallback_status(ok: bool, ok_text: str, *, action_text: str = '',
                     row_marks: list[RowMark] | None = None,
                     error_detail: str = '') -> dict:
    """events 없이 ok 여부만 알 때 사용하는 최소 status.
    error_detail: logs에서 추출한 첫 오류 메시지 — UI 카드에 표시됨.
    """
    if ok:
        return build_status('ok', [], summary_text=ok_text,
                            detail_messages=[], action_text=action_text,
                            row_marks=_row_mark_summary(row_marks or []))
    msgs = [{'level': 'error', 'text': error_detail}] if error_detail else []
    return build_status(
        'error', msgs,
        summary_text='오류가 있습니다.',
        detail_messages=[error_detail] if error_detail else [],
        action_text=action_text,
        row_marks=_row_mark_summary(row_marks or []),
    )


def normalize_scan_item(meta, kind_label, file_key=None, events=None):
    if meta is None:
        return None
    data_start_row = meta_get(meta, 'data_start_row', 2)
    issue_rows = as_abs_issue_rows(meta_get(meta, 'issue_rows', []) or [], data_start_row)
    file_evts = [e for e in (events or []) if getattr(e, 'file_key', None) == file_key] if file_key else []
    if file_evts:
        errs = [e for e in file_evts if e.level == 'error']
        warns = [e for e in file_evts if e.level == 'warn']
        if errs:
            severity = 'error'
            warning = errs[0].message
        elif warns:
            severity = 'warn'
            warning = '\n'.join(e.message for e in warns)
        else:
            severity = 'ok'
            warning = ''
    else:
        severity = meta_get(meta, 'severity', 'ok')
        warning = meta_get(meta, 'warning', '')

    item = {
        'kind': kind_label,
        'file_name': meta_get(meta, 'file_name', ''),
        'file_path': meta_get(meta, 'file_path', ''),
        'sheet_name': meta_get(meta, 'sheet_name', ''),
        'header_row': meta_get(meta, 'header_row', 1),
        'data_start_row': data_start_row,
        'row_count': meta_get(meta, 'row_count', 0),
        'warning': warning,
        'issue_rows': issue_rows,
        'extra_grades': list(meta_get(meta, 'extra_grades', []) or []),
        'severity': severity,
        'messages': [{'level': e.level, 'text': e.message} for e in file_evts],
    }
    item['status'] = status_from_events(file_evts) if file_evts else build_status(
        severity if severity in ('error', 'warn') else 'ok',
        [],
        summary_text='확인 완료',
        detail_messages=[],
        row_marks={'warn_rows': [], 'error_rows': [], 'issue_rows': issue_rows},
    )
    return item


def normalize_compare_item(compare_file, compare_layout):
    if not compare_file:
        return None
    compare_layout = compare_layout or {}
    issue_rows = as_abs_issue_rows(meta_get(compare_layout, 'issue_rows', []) or [], meta_get(compare_layout, 'data_start_row', 2))
    warning = meta_get(compare_layout, 'warning', '')
    item = {
        'kind': '재학생',
        'file_name': getattr(compare_file, 'name', str(compare_file)),
        'file_path': str(compare_file),
        'sheet_name': meta_get(compare_layout, 'sheet_name', ''),
        'header_row': meta_get(compare_layout, 'header_row', 1),
        'data_start_row': meta_get(compare_layout, 'data_start_row', 2),
        'row_count': meta_get(compare_layout, 'row_count', 0),
        'warning': warning,
        'issue_rows': issue_rows,
        'extra_grades': [],
        'severity': meta_get(compare_layout, 'severity', 'warn' if warning or issue_rows else 'ok'),
    }
    item['status'] = build_status(
        item['severity'] if item['severity'] in ('error', 'warn') else 'ok',
        [],
        summary_text='확인 완료',
        detail_messages=[],
        row_marks={'warn_rows': issue_rows, 'error_rows': [], 'issue_rows': issue_rows},
    )
    return item


def _grade_year_map_from_scan_result(result) -> dict:
    grade_year_map = {}
    roster_info = getattr(result, 'roster_info', None)
    if roster_info is None:
        return grade_year_map

    target_grades = set()
    prefix_mode = getattr(roster_info, 'prefix_mode_by_roster_grade', {}) or {}
    shift = int(getattr(roster_info, 'ref_grade_shift', 0) or 0)
    for g_roster in prefix_mode.keys():
        try:
            g_cur = int(g_roster) - shift
        except Exception:
            continue
        if g_cur > 0:
            target_grades.add(g_cur)
    freshmen_meta = getattr(result, 'freshmen', None) or {}
    for g in (meta_get(freshmen_meta, 'extra_grades', []) or []):
        try:
            g_i = int(g)
        except Exception:
            continue
        if g_i > 0:
            target_grades.add(g_i)
    target_grades.update({1, 2, 3, 4, 5, 6})
    year_int = int(getattr(result, 'year_int', 0) or 0)
    derived = derive_grade_year_map(target_grades=sorted(target_grades), input_year=year_int, roster_info=roster_info)
    return {int(k): int(v) for k, v in derived.items()}


def present_scan_result(result) -> dict:
    logs = logs_from_result(result)
    events = list(getattr(result, 'events', None) or [])
    row_marks = list(getattr(result, 'row_marks', None) or [])

    items = [
        i for i in [
            normalize_scan_item(getattr(result, 'freshmen', None), '신입생', 'freshmen', events),
            normalize_scan_item(getattr(result, 'transfer_in', None), '전입생', 'transfer_in', events),
            normalize_scan_item(getattr(result, 'transfer_out', None), '전출생', 'transfer_out', events),
            normalize_scan_item(getattr(result, 'teachers', None), '교직원', 'teachers', events),
            normalize_compare_item(getattr(result, 'compare_file', None), getattr(result, 'compare_layout', None)),
        ] if i is not None
    ]

    rbd = getattr(result, 'roster_basis_date', None)
    roster_basis_date = rbd.isoformat() if hasattr(rbd, 'isoformat') else (str(rbd) if rbd is not None else '')

    _first_err = next((l['message'] for l in logs if l['level'] == 'error'), '')
    status = status_from_events(events) if events else _fallback_status(
        bool(getattr(result, 'ok', False)), '스캔 완료', row_marks=row_marks, error_detail=_first_err
    )
    if events:
        status['row_marks'] = _row_mark_summary(row_marks)

    return {
        'ok': bool(getattr(result, 'ok', False)),
        'school_profile_mode': str(getattr(result, 'school_profile_mode', 'single') or 'single'),
        'school_kind_needs_choice': bool(getattr(result, 'school_kind_needs_choice', False)),
        'grade_rule_max_grade': int(getattr(result, 'grade_rule_max_grade', 6) or 6),
        'can_execute': bool(getattr(result, 'can_execute', False)),
        'can_execute_after_input': bool(getattr(result, 'can_execute_after_input', False)),
        'missing_fields': list(getattr(result, 'missing_fields', []) or []),
        'needs_open_date': bool(getattr(result, 'needs_open_date', False)),
        'need_roster': bool(getattr(result, 'need_roster', False)),
        'roster_date_mismatch': bool(getattr(result, 'roster_date_mismatch', False)),
        'roster_basis_date': roster_basis_date,
        'roster_path': str(getattr(result, 'roster_path', None) or '') or None,
        'has_school_kind_warn': False,
        'grade_year_map': _grade_year_map_from_scan_result(result),
        'items': items,
        'logs': logs,
        'status': status,
        'events': [serialize_event(e) for e in events],
        'row_marks': [serialize_row_mark(m) for m in row_marks],
    }


def present_run_result(result) -> dict:
    logs = logs_from_result(result)
    events = list(getattr(result, 'events', None) or [])
    row_marks = list(getattr(result, 'row_marks', None) or [])
    audit = getattr(result, 'audit_summary', {}) or {}
    in_cnt = audit.get('input_counts', {})
    action_text = ''
    if any('헤더를 찾을 수 없습니다.' in str(l.get('message', '')) for l in logs):
        action_text = '헤더행과 열 이름을 확인해 주세요.'

    _first_err = next((l['message'] for l in logs if l['level'] == 'error'), '')
    status = status_from_events(events) if events else _fallback_status(
        bool(getattr(result, 'ok', False)), '완료', action_text=action_text, row_marks=row_marks, error_detail=_first_err
    )

    return {
        'ok': bool(getattr(result, 'ok', False)),
        'output_files': [{'name': p.name, 'path': str(p)} for p in (getattr(result, 'outputs', None) or [])],
        'freshmen_count': int(in_cnt.get('freshmen', 0)),
        'teacher_count': int(in_cnt.get('teacher', 0)),
        'transfer_in_done': int(getattr(result, 'transfer_in_done', 0)),
        'transfer_in_hold': int(getattr(result, 'transfer_in_hold', 0)),
        'transfer_out_done': int(getattr(result, 'transfer_out_done', 0)),
        'transfer_out_hold': int(getattr(result, 'transfer_out_hold', 0)),
        'transfer_out_auto_skip': int(getattr(result, 'transfer_out_auto_skip', 0)),
        'notice_dup_rows': list(getattr(result, 'notice_dup_rows', []) or []),
        'notice_teacher_dup_rows': list(getattr(result, 'notice_teacher_dup_rows', []) or []),
        'logs': logs,
        'status': status,
        'events': [serialize_event(e) for e in events],
        'row_marks': [serialize_row_mark(m) for m in row_marks],
    }


def present_diff_run_result(result) -> dict:
    logs = logs_from_result(result)
    events = list(getattr(result, 'events', None) or [])
    row_marks = list(getattr(result, 'row_marks', None) or [])
    _first_err = next((l['message'] for l in logs if l['level'] == 'error'), '')
    status = status_from_events(events) if events else _fallback_status(
        bool(getattr(result, 'ok', False)), '완료', row_marks=row_marks, error_detail=_first_err
    )
    return {
        'ok': bool(getattr(result, 'ok', False)),
        'output_files': [{'name': p.name, 'path': str(p)} for p in (getattr(result, 'outputs', None) or [])],
        'compare_only_count': int(getattr(result, 'compare_only_count', 0)),
        'roster_only_count': int(getattr(result, 'roster_only_count', 0)),
        'matched_count': int(getattr(result, 'matched_count', 0)),
        'unresolved_count': int(getattr(result, 'unresolved_count', 0)),
        'transfer_in_done': int(getattr(result, 'transfer_in_done', 0)),
        'transfer_in_hold': int(getattr(result, 'transfer_in_hold', 0)),
        'transfer_out_done': int(getattr(result, 'transfer_out_done', 0)),
        'transfer_out_hold': int(getattr(result, 'transfer_out_hold', 0)),
        'roster_only_rows': list(getattr(result, 'roster_only_rows', []) or []),
        'matched_rows': list(getattr(result, 'matched_rows', []) or []),
        'compare_only_rows': list(getattr(result, 'compare_only_rows', []) or []),
        'unresolved_rows': list(getattr(result, 'unresolved_rows', []) or []),
        'logs': logs,
        'status': status,
        'events': [serialize_event(e) for e in events],
        'row_marks': [serialize_row_mark(m) for m in row_marks],
    }
