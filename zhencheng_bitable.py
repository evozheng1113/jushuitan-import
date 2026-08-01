"""真诚部门 - 天然钻工厂单 → 飞书多维表 镶嵌成本同步 (v23.1)

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
import re
from feishu_client import FeishuClient, load_credentials
from natural import _norm_code


# ============ 真诚部门多维表定位 ============
ZHENCHENG_APP_TOKEN = 'FNtHbYiW3a4dJasJJbucWlwMnlc'
ZHENCHENG_TABLE_ID  = 'tblzKxchN598phDb'


# ============ 字段名候选 (运行时自适应, 命中第一个就用) ============
CERT_FIELD_CANDIDATES = ['证书编码', '证书编号', '证书号', '商品编码']
NAME_FIELD_CANDIDATES = ['客户名', '客户名称', '客户', '下单单号', '款号']
COST_FIELD_CANDIDATES = ['镶嵌成本', '成本2', '镶嵌费', '加工费', '镶工费', '工费']


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


def _pick_field(fields_meta, candidates):
    names = {f.get('field_name', '').strip() for f in fields_meta}
    for c in candidates:
        if c in names:
            return c
    return None


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
        'details': [], 'dry_run': dry_run,
    }

    # ---- 探测字段 ----
    fields_meta = client.list_fields(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID)
    if not fields_meta:
        raise RuntimeError("真诚多维表 list_fields 返回空, 检查 app_token/table_id/权限")

    cert_field = _pick_field(fields_meta, CERT_FIELD_CANDIDATES)
    name_field = _pick_field(fields_meta, NAME_FIELD_CANDIDATES)
    cost_field = _pick_field(fields_meta, COST_FIELD_CANDIDATES)

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

    result['cert_field'] = cert_field
    result['name_field'] = name_field
    result['cost_field'] = cost_field

    # ---- 逐条同步 ----
    for it in items:
        cost = it.get('镶嵌成本')
        cert = str(it.get('证书号') or '').strip()
        name = str(it.get('款号') or '').strip()
        no   = it.get('no')

        if not cost:
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'skip', 'reason': '成本为空'})
            continue
        if not (cert or name):
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'skip', 'reason': '证书 & 款号都空'})
            continue

        rec = None
        matched_by = None

        # 1. 证书号优先
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

        # 2. 兜底: 款号 (客户名)
        if not rec and name_field and name:
            try:
                r = client.find_by_field(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                          name_field, name)
                if r:
                    rec, matched_by = r, f'{name_field}={name}'
            except Exception as e:
                result['errors'].append(f"#{no} 款号查询 '{name}' 失败: {e}")

        if not rec:
            desc = f"cert={cert or '空'} name={name or '空'}"
            result['not_found'].append(desc)
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'not_found',
                                       'reason': '证书 & 款号都没匹配到'})
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
            client.update_record(ZHENCHENG_APP_TOKEN, ZHENCHENG_TABLE_ID,
                                 rec_id, {cost_field: cost})
            result['updated'] += 1
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'updated',
                                       'matched_by': matched_by, 'cost': cost})
        except Exception as e:
            result['errors'].append(f"#{no} 写入失败: {e}")
            result['details'].append({'no': no, 'cert': cert, 'name': name,
                                       'status': 'error', 'reason': str(e)})

    return result


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
