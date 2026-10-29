"""布心精确匹配与持久同步记录。不依赖文件名、行号或服务器本地缓存。"""
import hashlib
import json
import re
from collections import defaultdict
import requests

MARKER_FIELD = '布心成本同步记录'


def _text(value):
    if isinstance(value, list):
        return ''.join(str(v.get('text', '')) if isinstance(v, dict) else str(v) for v in value)
    return str(value or '').strip()


def customer_candidates(raw):
    value = str(raw or '').strip()
    if '-' in value:
        value = value.split('-', 1)[1].strip()
    else:
        match = re.match(r'^[\u4e00-\u9fff]{2,4}\s+(.+)$', value)
        if match:
            value = match.group(1).strip()
        else:
            # 无分隔符时不猜测销售姓名长度；保留整串供精确匹配。
            return [value] if value else []
    clean = re.sub(r'(?:\d{1,2}[./]\d{1,2}(?:-\d+)?|\d{6})$', '', value).strip()
    return list(dict.fromkeys(v for v in (value, clean) if v))


def _search(client, app, table, field, values, factory=None):
    records = {}
    for value in values:
        conditions = [{'field_name': field, 'operator': 'is', 'value': [value]}]
        if factory:
            conditions.append({'field_name': factory, 'operator': 'is', 'value': ['布心']})
        response = requests.post(
            f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app}/tables/{table}/records/search',
            headers=client._headers(), json={'filter': {'conjunction': 'and', 'conditions': conditions}}, timeout=15)
        data = response.json()
        if data.get('code') != 0:
            raise RuntimeError(f"查询失败：{data.get('msg') or data.get('code')}")
        page = data.get('data') or {}
        found = page.get('items') or []
        if page.get('has_more') or len(found) > 1:
            raise ValueError(f'{field}={value} 对应多张订单，请补充单号或证书')
        for rec in found:
            if not rec.get('record_id'):
                raise ValueError('查询结果缺少记录编号')
            records[rec['record_id']] = (rec, f'{field}={value}')
    if len(records) > 1:
        raise ValueError('候选信息对应不同订单，请核对')
    return next(iter(records.values()), (None, None))


def prepare_matches(client, items, fields, app, table):
    names = {f.get('field_name') for f in fields}
    def pick(options):
        return next((n for n in options if n in names), None)
    order_field = pick(['下单单号'])
    factory_field = pick(['下单工厂'])
    cert_field = pick(['证书编码', '证书编号', '证书号'])
    name_field = pick(['客户名称', '客户名', '客户'])
    result, by_record = {}, defaultdict(list)
    for item in items:
        if item.get('_工厂') != '布心':
            continue
        try:
            order = str(item.get('下单单号') or '').strip()
            cert = str(item.get('证书号') or '').strip()
            if order:
                if not order_field:
                    raise ValueError('飞书缺少下单单号字段')
                is_f = bool(re.fullmatch(r'F-\d{1,2}-\d{1,2}-\d+', order, re.I))
                bare = bool(re.fullmatch(r'\d{1,2}-\d{1,2}-\d+', order))
                if not is_f and not bare:
                    raise ValueError(f'订单号 {order} 不是 F 编号或过渡期日期编号，请核对')
                if bare and not factory_field:
                    raise ValueError('无 F 编号必须核对飞书下单工厂，当前缺少该字段')
                rec, via = _search(client, app, table, order_field, [order], factory_field if bare else None)
                if not rec:
                    raise ValueError(f'未找到订单 {order}' + (' 且下单工厂为布心的记录' if bare else ''))
            else:
                rec, via = None, None
                if cert and cert_field:
                    variants = [cert]
                    if cert.endswith('-1'): variants.append(cert[:-2])
                    else: variants.append(cert + '-1')
                    rec, via = _search(client, app, table, cert_field, variants, factory_field)
                if not rec and name_field:
                    if not factory_field:
                        raise ValueError('客户名匹配需核对下单工厂，当前缺少该字段')
                    rec, via = _search(client, app, table, name_field,
                                       customer_candidates(item.get('款号')), factory_field)
                if not rec:
                    raise ValueError('未找到唯一订单，请补充下单单号或证书；不会用销售姓名匹配')
            result[id(item)] = (rec, via, None)
            by_record[rec['record_id']].append(id(item))
        except Exception as exc:
            result[id(item)] = (None, None, str(exc))
    for ids in by_record.values():
        if len(ids) > 1:
            for key in ids:
                result[key] = (None, None, '本次多件出货匹配同一订单，请核对单号或证书后再同步')
    return result


def prepare_ledger(item, rec, fields, cost_field):
    field = next((f for f in fields if f.get('field_name') == MARKER_FIELD), None)
    if not field or field.get('type') != 1:
        raise ValueError(f'请先在真诚客户多维表新增文本字段「{MARKER_FIELD}」')
    date = item.get('_布心批次')
    if not date or not re.fullmatch(r'20\d{2}-\d{2}-\d{2}', date):
        raise ValueError('未识别出货批次日期，暂不写入')
    source = item.get('_布心原始')
    if not isinstance(source, dict) or not source:
        raise ValueError('缺少原始成本信息，暂不写入')
    existing = rec.get('fields') or {}
    raw = _text(existing.get(MARKER_FIELD))
    if raw:
        try:
            ledger = json.loads(raw)
            if ledger.get('version') != 1 or not isinstance(ledger.get('shipments'), dict):
                raise ValueError()
        except (ValueError, TypeError, AttributeError):
            raise ValueError('原同步记录格式异常，请核对，暂不覆盖')
    else:
        ledger = {'version': 1, 'shipments': {}}
    # 飞书 record_id 提供订单身份；同一订单同材质同日多件由预检查拦截。
    key = '布心|' + item['材质颜色'] + '|' + date
    payload = {'source': source, 'cost': item['镶嵌成本'],
               'order': item.get('下单单号'), 'cert': item.get('证书号'), 'customer': item.get('款号')}
    fingerprint = hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    previous = ledger['shipments'].get(key)
    if previous:
        if not isinstance(previous, dict) or previous.get('fingerprint') != fingerprint:
            raise ValueError('这批已同步，但重量、费用、金价或订单信息有变化，请人工核对')
        return None, True
    current_cost = existing.get(cost_field)
    if not raw and current_cost not in (None, '', 0, 0.0) and not item.get('_布心首次确认'):
        raise ValueError('飞书已有成本但没有同步记录；请核对该材质最后一批是否已处理')
    ledger['shipments'][key] = {'fingerprint': fingerprint, 'cost': item['镶嵌成本']}
    encoded = json.dumps(ledger, ensure_ascii=False, separators=(',', ':'))
    if len(encoded) > 18000:
        raise ValueError('同步记录已满，请先归档后再处理')
    return encoded, False
