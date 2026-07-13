"""成品新单同步 (v20)
================================
把工厂现货 (XH 单号) 或 客户单 (证书号兜底) 的镶嵌成本
同步到飞书电子表格 "成品新单":
- 匹配: sheet 里 P 列 (第 16 列) = 我们的 match_key
- 写入: 匹配行的 M 列 (第 13 列) 叠加镶嵌成本

用户飞书表格: https://icnikyg9yylq.feishu.cn/sheets/IIwus3ROBhAR1et2anJcvdkTnNf?sheet=1XBqwN
"""
import math
import requests

FEISHU_BASE = 'https://open.feishu.cn/open-apis'
# 默认成品新单表 (可通过 client 构造函数覆盖)
FP_TOKEN = 'IIwus3ROBhAR1et2anJcvdkTnNf'
FP_SHEET_ID = '1XBqwN'   # sheet2 = 成品新单

# 列位 (0-based 索引)
COL_H = 7    # H 列 → 工厂状态 ("二厂待出货" → 匹配上同步后改为"二厂")
COL_M = 12   # M 列 → 镶嵌成本 (叠加写入, 固定)
# v22.5: 不再硬编码匹配键的列位, 每行扫全部单元格自动识别
#        (证书号可能在 G 列, XH 单号可能在 P 列, 各表不同. 让代码自己判断)
COL_P = 6    # 兼容旧引用


class FinishedProductsClient:
    """飞书电子表格客户端 (只针对"成品新单"表, 支持读+批量写 M 列)"""

    def __init__(self, app_id, app_secret,
                 spreadsheet_token=FP_TOKEN, sheet_id=FP_SHEET_ID):
        import time as _time_mod
        self.app_id = app_id
        self.app_secret = app_secret
        self.spreadsheet_token = spreadsheet_token
        self.sheet_id = sheet_id
        self._tenant_token = None
        self._token_expire = 0.0
        self._time_mod = _time_mod

    def _get_tenant_token(self):
        """v20.2: token 2 小时过期, 加提前 5 分钟自动刷新"""
        now = self._time_mod.time()
        if self._tenant_token and self._token_expire > now:
            return self._tenant_token
        r = requests.post(
            f'{FEISHU_BASE}/auth/v3/tenant_access_token/internal',
            json={'app_id': self.app_id, 'app_secret': self.app_secret},
            timeout=10)
        d = r.json()
        if d.get('code') != 0:
            raise RuntimeError(f"飞书 token 失败: {d}")
        self._tenant_token = d['tenant_access_token']
        self._token_expire = now + d.get('expire', 7200) - 300
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

    def batch_write(self, updates):
        """updates: [(row_idx, col_letter, value), ...]
           行号 1-based (含表头行), col_letter 'M'/'H'/'A' 等
           支持一次写多列到多行
        """
        if not updates:
            return {'code': 0, 'updated_rows': 0, 'skipped': True}
        value_ranges = [
            {'range': f'{self.sheet_id}!{col}{r}:{col}{r}', 'values': [[v]]}
            for r, col, v in updates
        ]
        # 飞书批量: 每批最多 100 range
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

    # 兼容旧调用
    def batch_write_M(self, updates):
        return self.batch_write([(r, 'M', v) for r, v in updates])


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

    # v22.10: 队列消费 (支持同 key 多条: 广州1行=N件拆成 N 条同 key)
    #         猛哥场景不受影响, 因为 parse_E 已在解析层合并 rows+cost 相加,
    #         build_sync_items 生成的是 1 条 sync_item (队列 1 条)
    from collections import deque
    key_queue = {}   # match_key → deque of items
    for it in items:
        k = str(it.get('match_key') or '').strip()
        if not k:
            continue
        key_queue.setdefault(k, deque()).append(dict(it))

    updates = []       # [(row_idx, col_letter, value)]
    match_log = []
    all_keys_seen = set(key_queue.keys())

    def _find_matching_key(p_val):
        """v20.9: 支持成品新单 P 列尾部带 -N 后缀 (如 P=755505076-2 匹配 fly_key=755505076)"""
        if p_val in key_queue and key_queue[p_val]:
            return p_val
        for k, q in key_queue.items():
            if not k or not q or not p_val.startswith(k):
                continue
            rest = p_val[len(k):]
            import re as _re
            if _re.match(r'^-\d+$', rest):
                return k
        return None

    for row_idx, row in enumerate(rows, start=1):
        if row_idx == 1:
            continue  # 表头

        # v22.5: 每行扫全部单元格, 自动识别哪一列存了匹配键
        #        条件: 值含 '-' 且长度 >=5 (排除成本数字、单一字符等)
        matched_key = None
        matched_p_val = None
        for cell in row:
            s = str(cell or '').strip()
            if len(s) < 5 or '-' not in s:
                continue
            k = _find_matching_key(s)
            if k:
                matched_key = k
                matched_p_val = s
                break
        if not matched_key:
            continue

        # v22.10: 队列消费 (每次匹配 pop 一个, 同 key 多条时按顺序各消费一次)
        it = key_queue[matched_key].popleft()
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
        new_m = math.ceil(old_m + cost)
        updates.append((row_idx, 'M', new_m))

        # v20.1: H 列 "某某厂待出货" → "某某厂" (去掉"待出货")
        old_h = ''
        new_h = None
        if len(row) > COL_H:
            raw_h = row[COL_H]
            if raw_h not in (None, ''):
                old_h = str(raw_h).strip()
                if '待出货' in old_h:
                    new_h = old_h.replace('待出货', '').strip()
                    updates.append((row_idx, 'H', new_h))

        match_log.append({
            'name': it.get('name', ''),
            'match_key': matched_key,   # 工厂单 fly_key
            'p_val': matched_p_val,     # 成品新单里实际存的值 (含可能的 -N 后缀)
            'row': row_idx,
            'old_m': math.ceil(old_m),
            'add_cost': math.ceil(cost),
            'new_m': new_m,
            'old_h': old_h,
            'new_h': new_h,
        })
    # v22.10: remaining = 队列还有剩余的 items (工厂单里有但成品新单没足够对应行)
    remaining = []
    for k in sorted(key_queue.keys()):
        for it in key_queue[k]:
            remaining.append(f"{k}(¥{it.get('cost')} {it.get('name', '')})")

    # 批量写入
    response = None
    errors = []
    if updates:
        try:
            response = fp_client.batch_write(updates)
        except Exception as e:
            errors.append(f'写入失败: {e}')

    return {
        'matched': match_log,
        'unmatched': sorted(remaining),
        'errors': errors,
        'response': response,
    }


import re as _re_fp


def _split_multi(s):
    """把多单号/多证书号字符串按 换行/斜杠/中文逗号/顿号 拆开, 去空."""
    if not s:
        return []
    parts = _re_fp.split(r'[\n\r/、,，;；]+', str(s))
    return [p.strip() for p in parts if p and p.strip()]


def _strip_cert_prefix(s):
    """v22.8: 去证书号前缀 (工厂单常带 IGI/GIA/LG, 成品新单只存数字部分).
       例: 'IGI807626332' → '807626332'
           'GIA 3555452881' → '3555452881'
           'LG786602329' → '786602329'
    """
    if not s:
        return s
    m = _re_fp.match(r'^\s*(IGI|GIA|LG)\s*(.+)$', str(s), _re_fp.IGNORECASE)
    return m.group(2).strip() if m else str(s).strip()


def build_sync_items_from_factory_items(factory_items):
    """把 factories.parse_X 返回的 items 转成 sync_costs 需要的格式.
       v22.6: 只同步现货件 (客户单走飞书多维表).
       v22.7: 支持一件拆多个 sync_items (C 列多 XH 换行 / E 列多证书号换行),
              一件成本按拆分数量均分 (向上取整).

       优先级:
         1. 下单编号 C 列里含 XH 单号 → 每个 XH 一条
         2. 证书编号 E 列多证书号 → 每个证书号一条
         3. 兜底: 飞书匹配键 / 单号 / 证书号
    """
    SPOT_LIKE = ('现货', '部门-真诚')
    sync_items = []
    for it in factory_items:
        if it.get('类别') not in SPOT_LIKE:
            continue
        cost = it.get('镶嵌成本')
        if not cost:
            continue

        order = it.get('下单编号') or ''
        cert = it.get('证书编号') or ''
        invoice = it.get('单号') or ''

        # 优先 1: C 列拆出所有 A-XH-*/B-XH-*/D-XH-*/E-XH-* 类的单号
        keys = []
        for line in _split_multi(order):
            if _re_fp.search(r'[A-Za-z]-XH-', line):
                keys.append(line)

        # 优先 2: 若 C 列没有 XH, 从 E 列拆多证书号 (如两只戒指 2 个 IGI)
        # v22.8: 去 IGI/GIA/LG 前缀 (成品新单只存数字部分, 如 807626332-2)
        if not keys:
            for line in _split_multi(cert):
                if len(line) >= 5:
                    keys.append(_strip_cert_prefix(line))

        # 兜底: 单一 fly_key / D 列单号 / E 列证书号 (证书号也去前缀)
        if not keys:
            for candidate in (it.get('飞书匹配键'), invoice, cert):
                if candidate and str(candidate).strip():
                    val = str(candidate).strip()
                    keys.append(_strip_cert_prefix(val) if candidate is cert else val)
                    break

        if not keys:
            continue

        # v22.10: 广州场景 — 工厂单一行 = N 件同款 (数量列 > 1),
        #         需要拆成 N 条同 key sync_item, 每条 1/N 成本;
        #         成品新单 N 行同 XH → 每行分别匹配一条 (队列消费, sync_costs 里做)
        qty = it.get('件数') or 1
        try:
            qty = int(qty)
        except (ValueError, TypeError):
            qty = 1
        if qty < 1:
            qty = 1
        if len(keys) == 1 and qty > 1:
            keys = keys * qty   # 复制 N 份, 相同 match_key

        # 成本均分, 向上取整 (每只都 ceil, 总额可能略高于原成本 —— 偏保守)
        import math as _math_fp
        per_cost = _math_fp.ceil(cost / len(keys))
        pinming = it.get('品名') or ''
        no = it.get('no') or ''

        for idx, k in enumerate(keys, start=1):
            tag = f' ({idx}/{len(keys)})' if len(keys) > 1 else ''
            sync_items.append({
                'match_key': k,
                'cost': per_cost,
                'name': f'#{no} {pinming}{tag}',
                '_类别': it.get('类别'),
            })
    return sync_items
