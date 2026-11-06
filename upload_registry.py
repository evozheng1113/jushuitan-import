"""共享文件历史。只记录上传/生成结果，不把这些状态当作成本同步成功。"""
import hashlib
import io
import json
from datetime import datetime, timezone, timedelta
from urllib.parse import quote
import openpyxl
import requests

TITLE = '网站文件记录'
BASE = 'https://open.feishu.cn/open-apis'
HEADER = ['内容指纹', '文件名', '时间', '状态', '文件指纹']


def fingerprint(data):
    raw = hashlib.sha256(data).hexdigest()
    digest = hashlib.sha256(b'xlsx-content-v1\n')
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    try:
        for ws in wb.worksheets:
            digest.update(json.dumps(['sheet', ws.title], ensure_ascii=False).encode())
            for row_index, row in enumerate(ws.iter_rows(values_only=True), 1):
                values = list(row)
                while values and values[-1] is None: values.pop()
                if not values: continue
                digest.update(json.dumps([row_index, values], ensure_ascii=False,
                                         default=str, separators=(',', ':')).encode())
                digest.update(b'\n')
    finally:
        wb.close()
    return digest.hexdigest(), raw


def _request(client, method, suffix, **kwargs):
    response = getattr(requests, method)(BASE + suffix, headers=client._headers(json_body=True), timeout=30, **kwargs)
    data = response.json()
    if data.get('code') != 0:
        raise RuntimeError(f"飞书错误 {data.get('code')}：{str(data.get('msg') or '未知错误')[:250]}")
    return data.get('data') or {}


def _sheet(client):
    token = client.spreadsheet_token
    def lookup():
        data = _request(client, 'get', f'/sheets/v3/spreadsheets/{token}/sheets/query')
        matches = [s for s in data.get('sheets', []) if s.get('title') == TITLE]
        if len(matches) > 1: raise RuntimeError('文件记录工作表重名，请核对')
        return matches[0] if matches else None
    sheet = lookup()
    if not sheet:
        try:
            _request(client, 'post', f'/sheets/v2/spreadsheets/{token}/sheets_batch_update',
                     json={'requests': [{'addSheet': {'properties': {'title': TITLE}}}]})
        except Exception:
            # Another user may have created the same sheet concurrently.
            sheet = lookup()
            if not sheet: raise
        sheet = sheet or lookup()
        if not sheet: raise RuntimeError('未能创建文件记录工作表')
    return sheet


def _read(client, sheet):
    count = int((sheet.get('grid_properties') or {}).get('row_count') or 1000)
    if count > 100000: raise RuntimeError('文件记录过多，需归档后再核对')
    rows = []
    for start in range(1, count + 1, 5000):
        rng = f"{sheet['sheet_id']}!A{start}:E{min(count, start + 4999)}"
        data = _request(client, 'get', f'/sheets/v2/spreadsheets/{client.spreadsheet_token}/values/{quote(rng, safe="!:")}')
        rows.extend(data.get('valueRange', {}).get('values') or [])
    def cell_text(value):
        if value is None: return ''
        if isinstance(value, list):
            return ''.join(str(v.get('text', '')) if isinstance(v, dict) else str(v) for v in value)
        return str(value)
    rows = [[cell_text(v) for v in (row or [])] for row in rows]
    # 空表可能返回 []、[[]] 或多行 [null, ...]；这些都不是错误表头。
    while rows and not any(v.strip() for v in rows[-1]): rows.pop()
    if rows:
        header = [v.strip() for v in rows[0]]
        while header and not header[-1]: header.pop()
        if header != HEADER:
            raise RuntimeError('网站文件记录表头不符，请核对；不会覆盖原内容')
    return rows


def _append(client, sheet, values):
    _request(client, 'post', f'/sheets/v2/spreadsheets/{client.spreadsheet_token}/values_append',
             params={'insertDataOption': 'INSERT_ROWS'},
             json={'valueRange': {'range': f"{sheet['sheet_id']}!A1:E{max(2, int((sheet.get('grid_properties') or {}).get('row_count') or 1000))}",
                                  'values': values}})


def _row(content, raw, filename, status):
    now = datetime.now(timezone(timedelta(hours=8))).isoformat(timespec='seconds')
    return [content, str(filename), now, status, raw]


def check_and_record(client, data, filename):
    try:
        content, raw = fingerprint(data)
    except Exception as exc:
        raise RuntimeError(f'读取出货文件失败：{type(exc).__name__}：{exc}') from exc
    try:
        sheet = _sheet(client)
    except Exception as exc:
        raise RuntimeError(f'访问文件记录工作表失败：{exc}') from exc
    try:
        rows = _read(client, sheet)
    except Exception as exc:
        raise RuntimeError(f'读取上传记录失败：{exc}') from exc
    matches = [r for r in rows[1:] if r and (r[0] == content or (len(r) > 4 and r[4] == raw))]
    info = {'seen': bool(matches), 'generated': any(len(r) > 3 and r[3] == '已生成结果' for r in matches),
            'content': content, 'raw': raw, 'sheet': sheet}
    info.update(action_summary(matches))
    if not matches:
        try:
            _append(client, sheet, ([] if rows else [HEADER]) + [_row(content, raw, filename, '已加载')])
        except Exception as exc:
            # 查询成功和写入历史失败是两种状态，不要把前者也标成未知。
            info['record_error'] = f'保存上传记录失败：{exc}'
    return info


def record_generated(client, info, filename):
    if info.get('unknown') or info.get('generated'):
        return
    if info.get('record_error'):
        # 上次写入可能未提交或响应超时；重读后再决定是否需要表头。
        sheet = _sheet(client)
        rows = _read(client, sheet)
        _append(client, sheet, ([] if rows else [HEADER]) + [_row(info['content'], info['raw'], filename, '已生成结果')])
        info.pop('record_error', None)
    else:
        _append(client, info['sheet'], [_row(info['content'], info['raw'], filename, '已生成结果')])
    info['generated'] = True


ACTION_PREFIX = '操作v2:'


def action_summary(matches):
    summary = {'jst_generated': False, 'jst_scopes': [], 'cost_action': None, 'archives': {},
               'legacy_result': any(len(r)>3 and r[3]=='已生成结果' for r in matches)}
    for row in matches:
        if len(row)<4 or not row[3].startswith(ACTION_PREFIX):
            continue
        try:
            event = json.loads(row[3][len(ACTION_PREFIX):])
        except (ValueError, TypeError):
            continue
        if not isinstance(event, dict):
            continue
        if event.get('kind') == 'jst_generated':
            summary['jst_generated'] = True
            scope = event.get('scope', '')
            if scope and scope not in summary['jst_scopes']:
                summary['jst_scopes'].append(scope)
        if event.get('kind') == 'cost_sync':
            summary['cost_action'] = event
        if event.get('kind') == 'archive' and event.get('factory'):
            summary['archives'][event['factory']] = event.get('state')
        if event.get('kind') == 'spot_sync':
            summary['spot_sync'] = True
    return summary


# 这些状态表示流水已写入或可能已写入；started 表示上次开始写入后没有结果。
ARCHIVE_BLOCKING = {'started', 'written', 'uncertain', 'layout_incomplete'}
ARCHIVE_LABELS = {'started': '结果未确认', 'uncertain': '结果未确认', 'written': '已写入',
                  'layout_incomplete': '已写入，格式未完成', 'failed': '写入失败，未写入'}


def archive_state(client, factory, batch):
    """每次写流水前重新读记录表，不用页面缓存；读不到就报错，由调用方放弃写入。
    按工厂+本批内容判断：文件改名、或老sheet有变动，只要这批没变就算同一批。"""
    if not batch:
        raise RuntimeError('本批内容指纹缺失，无法确认是否写过流水')
    state = None
    for row in _read(client, _sheet(client))[1:]:
        if len(row) < 4 or not row[3].startswith(ACTION_PREFIX):
            continue
        try:
            event = json.loads(row[3][len(ACTION_PREFIX):])
        except (ValueError, TypeError):
            continue
        if isinstance(event, dict) and event.get('kind') == 'archive' \
                and event.get('factory') == factory and event.get('batch') == batch:
            state = event.get('state')
    return state


def status_lines(info):
    legacy = info.get('legacy_result', False)
    if info.get('jst_generated'):
        jst = '已生成' + ('（'+'、'.join(info.get('jst_scopes', []))+'）' if info.get('jst_scopes') else '')
    else:
        jst = '旧记录未区分' if legacy else '无生成记录'
    event = info.get('cost_action')
    if not event:
        cost = '旧记录未区分' if legacy else '无同步记录'
    else:
        written, already, failed = (event.get(k, 0) for k in ('written', 'already', 'failed'))
        if not event.get('complete'):
            cost = '结果未确认（处理中或中断）'
        elif failed:
            cost = '部分成功' if written or already else '未获成功确认'
        elif written:
            cost = '全部成功'
        elif already:
            cost = '已同步，本次跳过'
        else:
            cost = '本次无可同步项'
        if written or already or failed:
            cost += f'（写入{written} · 已同步跳过{already} · 待核对{failed}）'
    archives = info.get('archives') or {}
    archive = '、'.join(f"{name}{ARCHIVE_LABELS.get(state, '状态未知')}" for name, state in archives.items())
    return [f'聚水潭入库 Excel：{jst}', f'飞书成本同步：{cost}', f"工厂流水：{archive or '无写入记录'}"]


def record_action(client, info, filename, event):
    if info.get('unknown') or not info.get('content'):
        raise RuntimeError('文件记录不可用，无法保存分项操作状态')
    status = ACTION_PREFIX + json.dumps(event, ensure_ascii=False, separators=(',', ':'))
    sheet = info.get('sheet')
    if info.get('record_error') or not sheet:
        sheet = _sheet(client)
        rows = _read(client, sheet)
        prefix = [] if rows else [HEADER]
    else:
        prefix = []
    _append(client, sheet, prefix + [_row(info['content'], info['raw'], filename, status)])
    info['sheet'] = sheet
    info.pop('record_error', None)
    if event['kind'] == 'jst_generated':
        info['jst_generated'] = True
        scope = event.get('scope', '')
        if scope and scope not in info.setdefault('jst_scopes', []):
            info['jst_scopes'].append(scope)
    elif event['kind'] == 'cost_sync':
        info['cost_action'] = dict(event)
    elif event['kind'] == 'spot_sync':
        info['spot_sync'] = True
    elif event['kind'] == 'archive':
        info.setdefault('archives', {})[event['factory']] = event['state']
        info['archive_count'] = info.get('archive_count', 0) + 1


class CostProgress:
    def __init__(self, run_id):
        self.run_id = run_id
        self.written = self.already = self.failed = 0

    def add(self, total, written=0, already=0):
        written = max(0, min(int(written), int(total)))
        already = max(0, min(int(already), int(total)-written))
        self.written += written
        self.already += already
        self.failed += max(0, int(total)-written-already)

    def event(self, complete=False):
        return {'kind':'cost_sync', 'run_id':self.run_id, 'written':self.written,
                'already':self.already, 'failed':self.failed, 'complete':bool(complete)}


def sheet_confirmed(result):
    """匹配日志不等于写入确认；检查真实API回执，而不是包装层的code。"""
    response = result.get('response') or {}
    replies = response.get('responses') or []
    return bool(replies) and response.get('code') == 0 and all(r.get('code') == 0 for r in replies)


def already_synced(result):
    return sum(d.get('status') == 'skip' and '已同步' in str(d.get('reason', ''))
               for d in result.get('details', []))
