"""成品新单同步 (v20)
================================
把工厂现货 (XH 单号) 或 客户单 (证书号兜底) 的镶嵌成本
同步到飞书电子表格 "成品新单":
- 匹配: sheet 里 P 列 (第 16 列) = 我们的 match_key
- 写入: 匹配行的 M 列 (第 13 列) 叠加镶嵌成本

用户飞书表格: https://icnikyg9yylq.feishu.cn/sheets/IIwus3ROBhAR1et2anJcvdkTnNf?sheet=1XBqwN
"""
import requests

FEISHU_BASE = 'https://open.feishu.cn/open-apis'
# 默认成品新单表 (可通过 client 构造函数覆盖)
FP_TOKEN = 'IIwus3ROBhAR1et2anJcvdkTnNf'
FP_SHEET_ID = '1XBqwN'   # sheet2 = 成品新单

# 列位 (0-based 索引)
COL_M = 12   # M 列 → 镶嵌成本 (叠加写入)
COL_P = 15   # P 列 → 匹配键 (存 XH 单号 或 证书编号)


class FinishedProductsClient:
    """飞书电子表格客户端 (只针对"成品新单"表, 支持读+批量写 M 列)"""

    def __init__(self, app_id, app_secret,
                 spreadsheet_token=FP_TOKEN, sheet_id=FP_SHEET_ID):
        self.app_id = app_id
        self.app_secret = app_secret
        self.spreadsheet_token = spreadsheet_token
        self.sheet_id = sheet_id
        self._tenant_token = None

    def _get_tenant_token(self):
        if self._tenant_token:
            return self._tenant_token
        r = requests.post(
            f'{FEISHU_BASE}/auth/v3/tenant_access_token/internal',
            json={'app_id': self.app_id, 'app_secret': self.app_secret},
            timeout=10)
        d = r.json()
        if d.get('code') != 0:
            raise RuntimeError(f"飞书 token 失败: {d}")
        self._tenant_token = d['tenant_access_token']
        return self._tenant_token

    def _headers(self, json_body=False):
        h = {'Authorization': f'Bearer {self._get_tenant_token()}'}
        if json_body:
            h['Content-Type'] = 'application/json'
        return h

    def load_all_rows(self, max_col='AZ', max_row=2000):
        """拉全 sheet (A1:AZ2000, 覆盖 M/P 列足够), 返回二维数组"""
        rng = f'{self.sheet_id}!A1:{max_col}{max_row}'
        r = requests.get(
            f'{FEISHU_BASE}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{rng}',
            headers=self._headers(), timeout=30)
        return r.json().get('data', {}).get('valueRange', {}).get('values', []) or []

    def batch_write_M(self, updates):
        """updates: [(row_idx, new_M_value), ...]  行号 1-based (含表头行)
           批量写入 M 列
        """
        if not updates:
            return {'code': 0, 'updated_rows': 0, 'skipped': True}
        value_ranges = [
            {'range': f'{self.sheet_id}!M{r}:M{r}', 'values': [[v]]}
            for r, v in updates
        ]
        # 飞书批量: 每批最多 100 range, 大批分批
        BATCH = 100
        responses = []
        for i in range(0, len(value_ranges), BATCH):
            batch = value_ranges[i:i + BATCH]
            r = requests.post(
                f'{FEISHU_BASE}/sheets/v2/spreadsheets/{self.spreadsheet_token}/values_batch_update',
                headers=self._headers(json_body=True),
                json={'valueRanges': batch},
                timeout=30)
            responses.append(r.json())
        return {'code': 0, 'updated_rows': len(updates), 'responses': responses}


def sync_costs(fp_client, items):
    """把 items 列表里的镶嵌成本同步到成品新单 M 列 (叠加).

    items: [{'match_key': 'B-XH-7-6-1', 'cost': 3140, 'name': '#1 LOOP项链'}, ...]

    匹配规则:
        - 成品新单 P 列 (第 16 列) 值 == match_key → 命中
        - 命中后读 M 列原值 (0 fallback), 加上 cost, 写回

    返回 {
        'matched': [{'name', 'match_key', 'row', 'old_m', 'add_cost', 'new_m'}, ...],
        'unmatched': ['B-XH-7-6-99', ...],
        'errors': [...],
        'response': 飞书返回,
    }
    """
    rows = fp_client.load_all_rows()
    if not rows or len(rows) < 2:
        return {
            'matched': [],
            'unmatched': [it['match_key'] for it in items if it.get('match_key')],
            'errors': [f'成品新单为空或只有表头 (拉到 {len(rows)} 行)'],
            'response': None,
        }

    # 建 key → item 映射 (跳过没 match_key 的)
    key_to_item = {}
    for it in items:
        k = str(it.get('match_key') or '').strip()
        if k:
            key_to_item[k] = it

    updates = []       # [(row_idx, new_M)]
    match_log = []
    remaining = set(key_to_item.keys())

    for row_idx, row in enumerate(rows, start=1):
        if row_idx == 1:
            continue  # 表头
        if len(row) <= COL_P:
            continue
        p_val = str(row[COL_P] or '').strip()
        if not p_val or p_val not in key_to_item:
            continue

        it = key_to_item[p_val]
        # 读 M 列原值
        old_m = 0
        if len(row) > COL_M:
            raw = row[COL_M]
            if raw not in (None, ''):
                try:
                    old_m = float(raw)
                except (ValueError, TypeError):
                    old_m = 0
        cost = it.get('cost') or 0
        new_m = round(old_m + cost)

        updates.append((row_idx, new_m))
        match_log.append({
            'name': it.get('name', ''),
            'match_key': p_val,
            'row': row_idx,
            'old_m': round(old_m),
            'add_cost': round(cost),
            'new_m': new_m,
        })
        remaining.discard(p_val)

    # 批量写入
    response = None
    errors = []
    if updates:
        try:
            response = fp_client.batch_write_M(updates)
        except Exception as e:
            errors.append(f'写入失败: {e}')

    return {
        'matched': match_log,
        'unmatched': sorted(remaining),
        'errors': errors,
        'response': response,
    }


def build_sync_items_from_factory_items(factory_items):
    """把 factories.parse_X 返回的 items 转成 sync_costs 需要的格式.
       - 现货 XH 单号: match_key = 飞书匹配键 (完整单号 B-XH-*)
       - 客户单: match_key = 飞书匹配键 (客户单号 / 客户名)
       - 兜底: 有证书编号且没 match_key → match_key = 证书编号
    """
    sync_items = []
    for it in factory_items:
        cost = it.get('镶嵌成本')
        if not cost:
            continue
        match_key = it.get('飞书匹配键')
        if not match_key:
            cert = it.get('证书编号')
            if cert:
                match_key = str(cert).strip()
        if not match_key:
            continue
        pinming = it.get('品名') or ''
        no = it.get('no') or ''
        sync_items.append({
            'match_key': match_key,
            'cost': cost,
            'name': f'#{no} {pinming}',
            '_类别': it.get('类别'),
        })
    return sync_items
