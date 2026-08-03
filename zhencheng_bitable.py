"""真诚部门 - 天然钻工厂单 → 飞书多维表 镶嵌成本同步 (v24.1)

匹配逻辑 = 真诚聚水潭 build_jst_row 的对齐版本:
  聚水潭里 `商品编码 = natural._norm_code(证书号)` (去多杠, 补 -1 后缀).
  真诚多维表就是聚水潭上游的数据表, 记录的证书 & 客户名 = 同一批客户.

匹配优先级 (证书优先, 每件唯一; 款号=客户名会重复所以兜底):
  1. 证书号 → 生成 4 种 variants 依次试 (原样 / -1 后缀 / 去 IGI-GIA-LG 前缀 / 去前缀+-1)
  2. 找不到 → 用款号 (客户名) 兜底
  3. 都失败 → 记 not_found

写入: 覆盖 (直接把该记录的 镶嵌成本 字段设为 item['镶嵌成本']).

同 4 家 parser (黛宝 / 布心 / 猛哥 / 二厂) 都用这套. 每件工厂单一 item = 一次写入;
猛哥同款号多行 (一对耳钉两只) 会写 2 次覆盖 → 需要多维表侧本身也是每只一条记录, 否则丢第一只.
后者是多维表设计问题, 不在同步侧解决 (跟聚水潭输出策略一致).
"""
import re, time
from feishu_client import FeishuClient, load_credentials
from natural import _norm_code


# ============ 真诚部门多维表定位 ============
ZHENCHENG_APP_TOKEN = 'FNtHbYiW3a4dJasJJbucWlwMnlc'
ZHENCHENG_TABLE_ID  = 'tblzKxchN598phDb'


# ============ 字段名候选 (运行时自适应, 命中第一个就用) ============
CERT_FIELD_CANDIDATES = ['证书编码', '证书编号', '证书号', '商品编码']
NAME_FIELD_CANDIDATES = ['客户名称', '客户名', '客户', '下单单号', '款号']  # v24.2 客户名称优先
COST_FIELD_CANDIDATES = ['镶嵌成本', '成本3', '镶嵌费', '加工费', '镶工费', '工费']  # v27: 镶嵌 = 聚水潭成本3
# v24.1: 回读用 (公式字段, 只读)
PROFIT_FIELD_CANDIDATES = ['利润', '利润额', '毛利']
RATE_FIELD_CANDIDATES   = ['利润率', '毛利率']
# v27: 反向补空 (聚水潭 → 飞书, 只在飞书该字段空时才写)
# v27.1: 用户业务表里 "裸钻成本" / "配石成本" 是公式字段, 真正可写的是 "XX 副本"
# 候选里把可写副本放前面, 且 _pick_field 会跳过 type=20 (公式) 的字段
BARE_COST_CANDIDATES    = ['裸钻成本 副本', '裸钻成本 副本2', '裸钻成本',
                            '成本1', '裸石成本', '裸钻']
SIDE_COST_CANDIDATES    = ['配石成本 副本', '配石成本 副本2', '配石成本',
                            '成本2', '散货成本', '副石成本']

# 飞书字段 type: 1=文本 2=数字 3=单选 4=多选 5=日期 7=复选框 11=人员
# 15=超链 17=附件 18=单向关联 19=Lookup 20=公式 21=双向关联 22=地理位置 23=群组
# 24=创建人 1005=创建时间 1006=修改时间 1004=修改人
_READONLY_FIELD_TYPES = {19, 20}   # Lookup 和 Formula 都不能写


# ============ 利润率阈值 (触发警示) ============
PROFIT_RATE_LOW  = 0.15
PROFIT_RATE_HIGH = 0.70


_STRIP_CERT_RE = re.compile(r'^\s*(IGI|GIA|LG)\s*(.+)$', re.IGNORECASE)


def _strip_cert_prefix(s):
    """去 IGI/GIA/LG 前缀 (跟 finished_products._strip_cert_prefix 同逻辑)."""
    if not s:
        return s
    m = _STRIP_CERT_RE.match(str(s))
    return m.group(2).strip() if m else str(s).strip()


def _cert_variants(cert):
    """生成一个证书号的所有可能存法 (按优先级):
       - _norm_code(原样): 跟聚水潭商品编码完全一致 (IGI... → IGI...-1)
       - 原样: 有些老数据可能没 -1 后缀
       - _norm_code(去前缀): 去 IGI/GIA/LG + -1 (807626332-1)
       - 去前缀原样: 807626332
    """
    if not cert:
        return []
    s = str(cert).strip()
    if not s:
        return []
    stripped = _strip_cert_prefix(s)
    seen, out = set(), []
    for v in (_norm_code(s), s, _norm_code(stripped), stripped):
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def _pick_field(fields_meta, candidates, exclude_readonly=False):
    """按候选顺序找命中的字段名.
    v27.1: exclude_readonly=True 时跳过公式/Lookup 字段 (写入类字段用).
    """
    lookup = {f.get('field_name', '').strip(): f for f in fields_meta}
    for c in candidates:
        f = lookup.get(c)
        if not f:
            continue
        if exclude_readonly and f.get('type') in _READONLY_FIELD_TYPES:
            continue
        return c
    return None


def _name_variants(name):
    """v26.1: 客户名可能是 '销售-客户' 或 '客户-销售' 或纯客户名 (布心场景).
       生成所有候选: 原样 + 拆 '-' 后每一段.
       多维表里存的通常是纯客户名, 所以拆分兜底能提升命中率.
    """
    if not name:
        return []
    s = str(name).strip()
    if not s:
        return []
    out = [s]
    if '-' in s:
        for part in s.split('-'):
            p = part.strip()
            if p and p not in out and len(p) >= 2:  # 长度 <2 防误匹 (单字)
                out.append(p)
    return out


def _get_record_by_id(client, app_token, table_id, record_id):
    """GET 单条记录 (用于轮询公式刷新)."""
    import requests
    url = (f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}'
           f'/tables/{table_id}/records/{record_id}')
    r = requests.get(url, headers=client._headers(), timeout=10)
    return r.json().get('data', {}).get('record', {})


def _wait_formula_refresh(client, record_id, cost_field, target_cost,
                           profit_field, rate_field, max_wait=6):
    """v24.1: 写完镶嵌成本后轮询等公式刷新, 返回 (profit, rate).
       判断: 记录里 cost_field 已刷新到 target_cost, 再等 0.4s 让下游公式算完.
    """
    deadline = time.time() + max_wait
    last_rec = None
    while time.time() < deadline:
        try:
            rec = _get_record_by_id(client, ZHENCHENG_APP_TOKEN,
                                     ZHENCHENG_TABLE_ID, record_id)
        except Exception:
            rec = None
        if rec:
            last_rec = rec
            fields = rec.get('fields', {})
            cur_cost = FeishuClient.get_number(fields.get(cost_field))
            if cur_cost is not None and abs(cur_cost - target_cost) < 0.5:
                time.sleep(0.4)
                try:
                    rec2 = _get_record_by_id(client, ZHENCHENG_APP_TOKEN,
                                              ZHENCHENG_TABLE_ID, record_id)
                    if rec2:
                        fields = rec2.get('fields', {})
                except Exception:
                    pass
                profit = FeishuClient.get_number(fields.get(profit_field)) if profit_field else None
                rate   = FeishuClient.get_number(fields.get(rate_field)) if rate_field else None
                return profit, rate, True
        time.sleep(0.5)
    # 超时: 兜底读一次现有值 (可能公式还没刷新)
    if last_rec:
        fields = last_rec.get('fields', {})
        profit = FeishuClient.get_number(fields.get(profit_field)) if profit_field else None
        rate   = FeishuClient.get_number(fields.get(rate_field)) if rate_field else None
        return profit, rate, False
    return None, None, False


def sync_zhencheng_costs(client, items, dry_run=False):
    """把 items 里每件的 镶嵌成本 写到真诚多维表对应记录.

    Args:
        client  : FeishuClient 实例
        items   : natural.py 输出的 items, 每条至少含 {款号, 证书号, 镶嵌成本}
        dry_run : True = 只查不写

    Returns:
        dict: matched/updated/not_found/errors/details + 命中的三个字段名
    """
    result = {
        'matched': 0, 'updated': 0,
        'not_found': [], 'errors': [],
        'cert_field': None, 'name_field': None, 'cost_field': None,
        'profit_field': None, 'rate_field': None,   # v24.1
        'bare_cost_field': None, 'side_cost_field': None,   # v27
        'details': [], 'dry_run': dry_run,
    }

    # ---- 探测字段 ----
    fields_meta = client.list_fields(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID)
    if not fields_meta:
        raise RuntimeError("真诚多维表 list_fields 返回空, 检查 app_token/table_id/权限")

    # cert / name 用于查询 (匹配), 可以是公式字段 (公式返回值也能作为 filter 依据)
    cert_field = _pick_field(fields_meta, CERT_FIELD_CANDIDATES)
    name_field = _pick_field(fields_meta, NAME_FIELD_CANDIDATES)
    # cost 是写入类, 必须非只读
    cost_field = _pick_field(fields_meta, COST_FIELD_CANDIDATES, exclude_readonly=True)

    all_names = [f.get('field_name', '') for f in fields_meta]

    if not cost_field:
        raise RuntimeError(
            f"没找到镶嵌成本字段. 候选: {COST_FIELD_CANDIDATES}\n"
            f"表实际字段: {all_names}\n"
            f"→ 请把真实字段名加进 zhencheng_bitable.COST_FIELD_CANDIDATES"
        )
    if not (cert_field or name_field):
        raise RuntimeError(
            f"没找到匹配字段 (证书 or 款号).\n"
            f"证书候选: {CERT_FIELD_CANDIDATES}\n"
            f"款号候选: {NAME_FIELD_CANDIDATES}\n"
            f"表实际字段: {all_names}\n"
            f"→ 请加进 CERT_FIELD_CANDIDATES 或 NAME_FIELD_CANDIDATES"
        )

    # v24.1: 回读用 (公式字段, 找不到不算错)
    profit_field = _pick_field(fields_meta, PROFIT_FIELD_CANDIDATES)
    rate_field   = _pick_field(fields_meta, RATE_FIELD_CANDIDATES)
    # v27: 反向补空用 (写入类, 排除公式/Lookup)
    bare_cost_field = _pick_field(fields_meta, BARE_COST_CANDIDATES, exclude_readonly=True)
    side_cost_field = _pick_field(fields_meta, SIDE_COST_CANDIDATES, exclude_readonly=True)

    result['cert_field'] = cert_field
    result['name_field'] = name_field
    result['cost_field'] = cost_field
    result['profit_field'] = profit_field
    result['rate_field']   = rate_field
    result['bare_cost_field'] = bare_cost_field
    result['side_cost_field'] = side_cost_field

    # ---- 逐条同步 ----
    for it in items:
        cost = it.get('镶嵌成本')
        # v25: 字段名兼容两套 —— 天然钻(natural.py) '证书号/款号', 培育钻(factories.py) '证书编号/下单编号'
        cert = str(it.get('证书号') or it.get('证书编号') or '').strip()
        # 客户名优先用 GIA/飞书查到的 (培育钻已经预填过 飞书客户名), 兜底用款号/下单编号
        gia_cust = str(it.get('飞书客户名') or '').strip()
        kh = str(it.get('款号') or it.get('下单编号') or it.get('单号') or '').strip()
        no  = it.get('no')

        if not cost:
            result['details'].append({'no': no, 'cert': cert, 'name': gia_cust or kh,
                                       'status': 'skip', 'reason': '成本为空'})
            continue
        if not (cert or gia_cust or kh):
            result['details'].append({'no': no, 'cert': cert, 'name': '',
                                       'status': 'skip', 'reason': '证书 & 客户名 & 款号都空'})
            continue

        rec = None
        matched_by = None

        # 1. 证书号优先 (每件唯一, 最强定位)
        if cert_field and cert:
            for cv in _cert_variants(cert):
                try:
                    r = client.find_by_field(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                              cert_field, cv)
                except Exception as e:
                    result['errors'].append(f"#{no} 证书查询 '{cv}' 失败: {e}")
                    continue
                if r:
                    rec, matched_by = r, f'{cert_field}={cv}'
                    break

        # 2. GIA 客户名 (从飞书 GIA 表锁定到的真实客户名)
        # v26.1: 复合名 "销售-客户" 场景 → 用 variants 拆 - 每段都试
        if not rec and name_field and gia_cust:
            for nv in _name_variants(gia_cust):
                try:
                    r = client.find_by_field(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                              name_field, nv)
                except Exception as e:
                    result['errors'].append(f"#{no} GIA客户名 '{nv}' 查询失败: {e}")
                    continue
                if r:
                    rec, matched_by = r, f'{name_field}(GIA)={nv}'
                    break

        # 3. 兜底: 工厂单款号 (猛哥 B 列本身就是客户名; 布心 B 列也可能是"销售-客户")
        if not rec and name_field and kh and kh != gia_cust:
            for nv in _name_variants(kh):
                try:
                    r = client.find_by_field(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                              name_field, nv)
                except Exception as e:
                    result['errors'].append(f"#{no} 款号 '{nv}' 查询失败: {e}")
                    continue
                if r:
                    rec, matched_by = r, f'{name_field}(款号)={nv}'
                    break

        display_name = gia_cust or kh
        if not rec:
            desc = f"cert={cert or '空'} GIA客户={gia_cust or '空'} 款号={kh or '空'}"
            result['not_found'].append(desc)
            result['details'].append({'no': no, 'cert': cert, 'name': display_name,
                                       'status': 'not_found',
                                       'reason': f'GIA客户名={gia_cust or "无"} 款号={kh or "无"} 都没查到'})
            continue

        result['matched'] += 1
        rec_id = rec.get('record_id')

        if dry_run:
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'would_update',
                                       'matched_by': matched_by,
                                       'record_id': rec_id, 'cost': cost})
            continue

        try:
            # v27: 反向补空 — 聚水潭有值 + 飞书该字段空 → 补写 (证书号/裸钻/配石)
            # 先取匹配到的 rec 里现有 fields, 判断哪些空
            existing = rec.get('fields', {}) or {}
            payload = {cost_field: cost}   # 镶嵌成本永远覆盖
            filled_extras = []

            # 证书号: it['证书号'] → 用 _norm_code 归一化 (跟聚水潭商品编码同格式)
            if cert_field and cert:
                cur_cert = FeishuClient.get_text(existing.get(cert_field))
                if not cur_cert or not str(cur_cert).strip():
                    payload[cert_field] = _norm_code(cert) or cert
                    filled_extras.append(f'证书={payload[cert_field]}')

            # 裸钻成本 (聚水潭 成本1) = gia['cost1']
            gia_info = it.get('_gia') or {}
            bare = gia_info.get('cost1')
            if bare_cost_field and isinstance(bare, (int, float)) and bare > 0:
                cur_bare = FeishuClient.get_number(existing.get(bare_cost_field))
                if cur_bare in (None, 0):
                    payload[bare_cost_field] = round(bare, 2)
                    filled_extras.append(f'裸钻={payload[bare_cost_field]}')

            # 配石成本 (聚水潭 成本2) = gia['cost2']
            side = gia_info.get('cost2')
            if side_cost_field and isinstance(side, (int, float)) and side > 0:
                cur_side = FeishuClient.get_number(existing.get(side_cost_field))
                if cur_side in (None, 0):
                    payload[side_cost_field] = round(side, 2)
                    filled_extras.append(f'配石={payload[side_cost_field]}')

            client.update_record(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                 rec_id, payload)
            result['updated'] += 1

            # v24.1: 轮询等公式刷新, 回读利润/利润率写回 item
            profit, rate, refreshed = _wait_formula_refresh(
                client, rec_id, cost_field, cost, profit_field, rate_field)
            it['飞书利润']   = profit
            it['飞书利润率'] = rate
            note = ' [⚠️公式刷新慢]' if not refreshed else ''

            # 阈值警示
            warn = ''
            if rate is not None:
                if rate < PROFIT_RATE_LOW:
                    warn = f' 🔴利润率低({rate*100:.1f}%)'
                elif rate > PROFIT_RATE_HIGH:
                    warn = f' 🟡利润率异常高({rate*100:.1f}%)'

            extras_str = (' 补:' + '/'.join(filled_extras)) if filled_extras else ''
            result['details'].append({'no': no, 'cert': cert, 'name': display_name,
                                       'status': 'updated',
                                       'matched_by': matched_by, 'cost': cost,
                                       'profit': profit, 'rate': rate,
                                       'note': note + warn + extras_str})
        except Exception as e:
            result['errors'].append(f"#{no} 写入失败: {e}")
            result['details'].append({'no': no, 'cert': cert, 'name': display_name,
                                       'status': 'error', 'reason': str(e)})

    return result


# ============ v24.4 网页诊断: 拉字段 + 前 3 条样本 ============
def probe_zhencheng_table(client, sample_size=3):
    """返回 dict: fields=[{name,type,ui_type}], samples=[{fields: {...}}]
    UI 可以调用这个然后展示, 帮用户看清楚字段名和类型.
    """
    import requests
    fields = client.list_fields(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID)
    field_summary = [
        {'name': f.get('field_name'), 'type': f.get('type'),
         'ui_type': f.get('ui_type')}
        for f in fields
    ]
    # 拉 sample_size 条最近记录看数据形状
    url = (f'https://open.feishu.cn/open-apis/bitable/v1/apps/{ZHENCHENG_APP_TOKEN}'
           f'/tables/{ZHENCHENG_TABLE_ID}/records?page_size={sample_size}')
    r = requests.get(url, headers=client._headers(), timeout=10)
    data = r.json()
    samples = data.get('data', {}).get('items', []) if data.get('code') == 0 else []
    return {'fields': field_summary, 'samples': samples,
            'list_code': data.get('code'), 'list_msg': data.get('msg')}


# ============ CLI 探测 (调试用: 打印字段结构) ============
if __name__ == '__main__':
    import sys, json
    app_id, app_secret = load_credentials()
    client = FeishuClient(app_id, app_secret)
    fields = client.list_fields(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID)
    print(f"真诚多维表共 {len(fields)} 个字段:")
    for f in fields:
        print(f"  {f.get('field_name'):20s}  type={f.get('type')}  ui={f.get('ui_type')}")
    if '--json' in sys.argv:
        print(json.dumps(fields, ensure_ascii=False, indent=2))
