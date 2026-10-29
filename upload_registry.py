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
        raise RuntimeError(f"文件记录服务不可用：{data.get('msg') or data.get('code')}")
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
    if rows and rows[0] and rows[0] != HEADER:
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
    content, raw = fingerprint(data)
    sheet = _sheet(client)
    rows = _read(client, sheet)
    matches = [r for r in rows[1:] if r and (r[0] == content or (len(r) > 4 and r[4] == raw))]
    info = {'seen': bool(matches), 'generated': any(len(r) > 3 and r[3] == '已生成结果' for r in matches),
            'content': content, 'raw': raw, 'sheet': sheet}
    if not matches:
        _append(client, sheet, ([] if rows else [HEADER]) + [_row(content, raw, filename, '已加载')])
    return info


def record_generated(client, info, filename):
    if info.get('unknown') or info.get('generated'):
        return
    _append(client, info['sheet'], [_row(info['content'], info['raw'], filename, '已生成结果')])
    info['generated'] = True
