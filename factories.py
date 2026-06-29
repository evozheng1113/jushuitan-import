"""4 家工厂解析逻辑 (网站精简版 v23.11)
只保留 parse_*, 删除 write_* / XML patch (网站不回写工厂表)
"""
import openpyxl, re, os, subprocess, tempfile
from datetime import datetime, date


def _is_silver(material):
    if not material:
        return False
    s = str(material).upper()
    return 'S925' in s or '925' in s or '银' in str(material)


def _num(v):
    return v if isinstance(v, (int, float)) else 0


def _normalize_order(s):
    if not s:
        return ''
    return re.sub(r'-+', '-', str(s).strip())


def _unwrap_datetime_order(v):
    if isinstance(v, (datetime, date)):
        y = v.year
        if 2000 <= y < 2050:
            return f"{y - 2000}-{v.month}-{v.day}"
    return v


def _ensure_xlsx(path):
    """.xls 自动转 .xlsx (网站环境若有 libreoffice 可用)
    没有 libreoffice 直接抛错, 让用户上传 .xlsx
    """
    if path.endswith('.xls') and not path.endswith('.xlsx'):
        out = path.rsplit('.', 1)[0] + '.xlsx'
        if not os.path.exists(out):
            try:
                subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx',
                                path, '--outdir', os.path.dirname(out) or '.'],
                               check=True, timeout=30)
            except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
                raise RuntimeError("无法处理 .xls 文件，请先用 Excel/WPS 另存为 .xlsx 再上传")
        return out
    return path


# ============ 工厂 A: 雅希 (广州厂) ============
def parse_A(excel_path, pt_price, au_price, sheet_name=None, **kw):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    items, current_rows = [], []
    for r in range(6, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, (int, float)):
            if current_rows:
                items.append(_aggregate_A(ws, current_rows, pt_price, au_price))
            current_rows = [r]
        elif current_rows:
            current_rows.append(r)
    if current_rows:
        items.append(_aggregate_A(ws, current_rows, pt_price, au_price))

    # 旧金回收 合并到对应客户单
    processed_recycles = set()
    recycle_by_order = {}
    for it in items:
        if it['类别'] == '旧金回收' and it.get('下单编号'):
            key = str(it['下单编号']).strip()
            if key:
                recycle_by_order.setdefault(key, []).append(it)
    for it in items:
        if it['类别'] == '客户单' and it.get('下单编号'):
            key = str(it['下单编号']).strip()
            cands = [r for r in recycle_by_order.get(key, []) if id(r) not in processed_recycles]
            if cands:
                rec = cands[0]
                it['镶嵌成本'] = round((it['镶嵌成本'] or 0) + (rec['镶嵌成本'] or 0))
                processed_recycles.add(id(rec))

    merged = []
    i = 0
    while i < len(items):
        cur = items[i]
        nxt = items[i + 1] if i + 1 < len(items) else None
        if (cur['类别'] == '客户单' and nxt and nxt['类别'] == '旧金回收'
                and id(nxt) not in processed_recycles):
            cur['镶嵌成本'] = round((cur['镶嵌成本'] or 0) + (nxt['镶嵌成本'] or 0))
            processed_recycles.add(id(nxt))
            merged.append(cur)
            merged.append(nxt)
            i += 2
        else:
            merged.append(cur)
            i += 1
    return merged


def _aggregate_A(ws, rows, pt_price, au_price):
    r0 = rows[0]
    no = int(ws.cell(row=r0, column=1).value)
    order = ws.cell(row=r0, column=3).value
    invoice = ws.cell(row=r0, column=4).value
    cert = ws.cell(row=r0, column=5).value
    style = ws.cell(row=r0, column=6).value or ''
    material = ws.cell(row=r0, column=7).value
    qty = ws.cell(row=r0, column=8).value or 1
    total_weight = ws.cell(row=r0, column=9).value

    sum_M, sum_N = 0.0, 0.0
    for r in rows:
        sum_M += _num(ws.cell(row=r, column=13).value)
        sum_N += _num(ws.cell(row=r, column=14).value)

    main_w = _num(ws.cell(row=r0, column=18).value)
    side_w = _num(ws.cell(row=r0, column=23).value)
    if not side_w and len(rows) > 1:
        for rr in rows[1:]:
            side_w += _num(ws.cell(row=rr, column=23).value)

    invoice_str, order_str = str(invoice or ''), str(order or '')
    style_str = str(style or '')
    if '修理' in style_str:
        cat = '修理'
    elif '退石' in style_str or '退钱' in style_str:
        cat = '退还'
    elif '旧金回收' in str(cert or ''):
        cat = '旧金回收'
    elif order_str.startswith('2026-'):
        cat = '现货'
    elif '真诚' in invoice_str or '真诚' in order_str:
        cat = '部门-真诚'
    elif not order_str and not str(cert or ''):
        cat = '内部-跳过'
    else:
        cat = '客户单'

    is_silver = material == 'S925'

    if material == 'PT950':
        gp = pt_price
    elif material in ('18K', '14K', '10K', '24K'):
        gp = au_price
    elif is_silver:
        gp = ws.cell(row=r0, column=15).value or 23
    else:
        gp = 0

    cost = 0
    if cat == '旧金回收':
        ab = ws.cell(row=r0, column=28).value
        if isinstance(ab, (int, float)) and ab < 0:
            cost = round(ab)
        else:
            gold_cost = sum_M * (pt_price or 0) + sum_N * (au_price or 0)
            s = g = l = z = o = 0
            for r in rows:
                s += _num(ws.cell(row=r, column=21).value)
                g += _num(ws.cell(row=r, column=24).value)
                l += _num(ws.cell(row=r, column=25).value)
                z += _num(ws.cell(row=r, column=26).value)
                o += _num(ws.cell(row=r, column=27).value)
            cost = round(gold_cost + s + g + l + z + o)
    elif cat in ('修理', '内部-跳过', '退还'):
        cost = 0
    elif is_silver:
        ab = ws.cell(row=r0, column=28).value
        if isinstance(ab, (int, float)):
            cost = round(ab)
        else:
            p = _num(ws.cell(row=r0, column=16).value)
            s = g = l = z = o = 0
            for r in rows:
                s += _num(ws.cell(row=r, column=21).value)
                g += _num(ws.cell(row=r, column=24).value)
                l += _num(ws.cell(row=r, column=25).value)
                z += _num(ws.cell(row=r, column=26).value)
                o += _num(ws.cell(row=r, column=27).value)
            cost = round(p + s + g + l + z + o)
    else:
        gold_cost = sum_M * (pt_price or 0) + sum_N * (au_price or 0)
        s = g = l = z = o = 0
        for r in rows:
            s += _num(ws.cell(row=r, column=21).value)
            g += _num(ws.cell(row=r, column=24).value)
            l += _num(ws.cell(row=r, column=25).value)
            z += _num(ws.cell(row=r, column=26).value)
            o += _num(ws.cell(row=r, column=27).value)
        cost = round(gold_cost + s + g + l + z + o)

    fly_key = f'A-{order_str.strip()}' if cat == '客户单' and order else None
    return {'no': no, '类别': cat, '下单编号': order, '单号': invoice, '证书编号': cert,
            '品名': style, '成色': material, '件数': qty, '镶嵌成本': cost,
            '总重': total_weight,
            '主石重量': main_w if main_w > 0 else None,
            '副石重量': side_w if side_w > 0 else None,
            '飞书匹配键': fly_key}


# ============ 工厂 B: 倾诚 ============
def parse_B(excel_path, sheet_name=None, **kw):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[-1]]
    REPAIR = ['换扣', '旧扣抵扣', '加链', '换链', '换细链', '换粗链']
    items = []
    for r in range(9, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if not isinstance(a, (int, float)):
            continue
        no = int(a)
        pinming = ws.cell(row=r, column=2).value
        invoice = ws.cell(row=r, column=3).value
        invoice = _unwrap_datetime_order(invoice)
        material = ws.cell(row=r, column=5).value
        qty = ws.cell(row=r, column=6).value
        total_weight = ws.cell(row=r, column=7).value
        total = ws.cell(row=r, column=28).value or 0
        if not pinming and not invoice and not material and total == 0:
            continue
        inv = str(invoice or '').strip()
        if any(k in str(pinming or '') for k in REPAIR):
            cat = '维修'
        elif not inv:
            cat = '现货'
        elif re.match(r'^\d+-\d+-\d+$', inv):
            cat = '客户单'
        elif re.match(r'^B-\d+-\d+-\d+$', inv):
            cat = '客户单'
        elif inv.lower().startswith('tb'):
            cat = '客户单'
        elif re.match(r'^(\d+分|\d+号|[小中大])$', inv):
            cat = '现货'
        elif inv in ('星河款戒指', '粉蓝宝', '小帆鱼', '红绳', '鱼美人'):
            cat = '现货'
        else:
            cat = '客户单'
        # 飞书匹配键
        fly_key = None
        if cat == '客户单':
            if inv.startswith('B-'):
                fly_key = inv
            elif re.match(r'^\d+-\d+-\d+$', inv):
                fly_key = f'B-{inv}'
            else:
                fly_key = inv
        items.append({'no': no, '类别': cat, '品名': pinming, '单号': invoice,
                      '成色': material, '件数': qty, '镶嵌成本': round(total),
                      '总重': total_weight,
                      '飞书匹配键': fly_key})
    return items


# ============ 工厂 D: 黛宝 ============
def parse_D(excel_path, pt_price, au_price, **kw):
    excel_path = _ensure_xlsx(excel_path)
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    items = []
    for r in range(6, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if not isinstance(a, (int, float)):
            continue
        no = int(a)
        kuanhao = ws.cell(row=r, column=3).value
        name = ws.cell(row=r, column=7).value
        material = ws.cell(row=r, column=8).value
        total_weight = ws.cell(row=r, column=10).value
        zhezu = ws.cell(row=r, column=15).value or 0
        AL = ws.cell(row=r, column=38).value or 0
        cert = ws.cell(row=r, column=4).value
        ks = str(kuanhao or '').strip()
        if ks == '成品款式' or not ks:
            cat, fly_key = '现货', None
        elif re.match(r'^\d+-\d+-\d+$', ks):
            cat, fly_key = '客户单', f'D-{ks}'
        else:
            cat, fly_key = '客户单', ks
        ms = str(material or '').upper()
        is_silver = _is_silver(material)
        if is_silver:
            cost = round(AL)
        else:
            if 'PT950' in ms:
                gp = pt_price
            else:
                gp = au_price
            cost = round(zhezu * gp + AL) if gp else round(AL)
        items.append({'no': no, '类别': cat, '款号': kuanhao,
                      '证书编号': cert, '品名': name, '成色': material,
                      '件数': 1, '镶嵌成本': cost,
                      '总重': total_weight,
                      '飞书匹配键': fly_key})
    return items


# ============ 工厂 E: 猛哥 ============
ORDER_LIKE = re.compile(r'^\d+-+\d+-+\d+$')
BATCH_LIKE = re.compile(r'^([A-Za-z]+\d+)[-—]?$')


def _detect_E_columns(ws):
    keys = ['序号', '单号', '编码', '品名', '成色', '手寸', '件数',
            '总重量', '净金重', '损耗', '含耗金重',
            '足金金重', '折足金', '金价', '金值',
            '配件重', '配件费', '主石', '副石',
            '镶主石费', '版蜡', '工艺费', '工费', '总金额']
    found = {}
    for c in range(1, 40):
        v = ws.cell(row=4, column=c).value
        if not v:
            continue
        s = str(v).strip()
        for key in keys:
            if key in s and key not in found:
                found[key] = c
                break
    if '折足金' not in found and '足金金重' in found:
        found['折足金'] = found['足金金重']
    if '品名' in found and ('成色' not in found or found['成色'] == found['品名']):
        next_col = found['品名'] + 1
        used = set(v for k, v in found.items() if k != '成色')
        if next_col not in used:
            found['成色'] = next_col
    if '件数' not in found and '成色' in found:
        found['件数'] = found['成色'] + 2
    if '总重量' not in found and '件数' in found:
        found['总重量'] = found['件数'] + 1
    if '主石' in found:
        found['主石重量'] = found['主石'] + 1
        found['主石金额'] = found['主石'] + 4
    if '副石' in found:
        found['副石重量'] = found['副石'] + 1
        found['副石金额'] = found['副石'] + 4
        found['镶石费'] = found['副石'] + 5
    if '工艺费' in found:
        found['版蜡'] = found['工艺费'] - 1
    defaults = {
        '序号': 1, '单号': 2, '品名': 3, '成色': 4, '件数': 6, '总重量': 7,
        '含耗金重': 10, '折足金': 11, '金价': 12, '金值': 13,
        '配件费': 15, '主石重量': 18, '主石金额': 20,
        '副石重量': 23, '副石金额': 26, '镶石费': 27,
        '镶主石费': 28, '版蜡': 29, '工艺费': 30, '工费': 31, '总金额': 32,
    }
    for k, v in defaults.items():
        if k not in found:
            found[k] = v
    return found


def parse_E(excel_path, pt_price, au_price, sheet_name=None, default_material=None, **kw):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[-1]]
    items = []
    auto_no_counter = 0
    COL = _detect_E_columns(ws)

    for r in range(6, ws.max_row + 1):
        a = ws.cell(row=r, column=COL['序号']).value
        b = ws.cell(row=r, column=COL['单号']).value
        b_str = str(b or '').strip()
        c_cell = ws.cell(row=r, column=COL['品名']).value
        d_cell = ws.cell(row=r, column=COL['成色']).value

        if not d_cell and default_material:
            d_cell = default_material

        no = None
        if isinstance(a, (int, float)):
            no = int(a)
        elif isinstance(a, str) and a.strip():
            astr = a.strip()
            stripped = astr.rstrip('-').rstrip('—')
            try:
                no = int(stripped)
            except ValueError:
                m = BATCH_LIKE.match(astr)
                if m:
                    no = m.group(1)
                elif stripped == '19楼':
                    no = '19楼'

        if no is None:
            if b_str and ORDER_LIKE.match(b_str):
                auto_no_counter += 1
                no = f"自动-{auto_no_counter}"
            elif c_cell and d_cell:
                auto_no_counter += 1
                no = f"现货-{auto_no_counter}"
            else:
                continue

        invoice = b
        pinming = c_cell
        material = d_cell
        qty = ws.cell(row=r, column=COL['件数']).value or 1
        total_weight = ws.cell(row=r, column=COL['总重量']).value
        han_hao = _num(ws.cell(row=r, column=COL['含耗金重']).value)
        factory_K = ws.cell(row=r, column=COL['折足金']).value
        factory_L = ws.cell(row=r, column=COL['金价']).value
        factory_M = ws.cell(row=r, column=COL['金值']).value
        peijian_fei  = _num(ws.cell(row=r, column=COL['配件费']).value)
        zhushi_jin   = _num(ws.cell(row=r, column=COL['主石金额']).value)
        fushi_jin    = _num(ws.cell(row=r, column=COL['副石金额']).value)
        xiangshi_fei = _num(ws.cell(row=r, column=COL['镶石费']).value)
        xzhushi_fei  = _num(ws.cell(row=r, column=COL['镶主石费']).value)
        banla        = _num(ws.cell(row=r, column=COL['版蜡']).value)
        gongyi       = _num(ws.cell(row=r, column=COL['工艺费']).value)
        gongfei      = _num(ws.cell(row=r, column=COL['工费']).value)
        main_w       = _num(ws.cell(row=r, column=COL['主石重量']).value)
        side_w       = _num(ws.cell(row=r, column=COL['副石重量']).value)

        if not invoice and not material:
            continue

        is_silver = _is_silver(material)
        mat_s = str(material or '').lower()

        if is_silver:
            ratio = 0
        elif 'pt' in mat_s:
            ratio = 0.952
        else:
            ratio = 0.755

        if isinstance(factory_K, (int, float)) and factory_K != 0:
            zhe_zu = factory_K
        elif han_hao and ratio:
            zhe_zu = round(han_hao * ratio, 5)
        else:
            zhe_zu = 0

        if isinstance(factory_L, (int, float)) and factory_L > 0:
            gp = factory_L
        elif 'pt' in mat_s:
            gp = pt_price or 0
        else:
            gp = au_price or 0

        if isinstance(factory_M, (int, float)) and factory_M > 0:
            jin_zhi = factory_M
        elif zhe_zu and gp:
            jin_zhi = round(zhe_zu * gp, 5)
        else:
            jin_zhi = 0

        if is_silver:
            factory_AF = _num(ws.cell(row=r, column=COL['总金额']).value)
            cost = round(factory_AF) if factory_AF else round(
                peijian_fei + zhushi_jin + fushi_jin + xiangshi_fei +
                xzhushi_fei + banla + gongyi + gongfei
            )
        else:
            cost = round(jin_zhi + peijian_fei + zhushi_jin + fushi_jin +
                         xiangshi_fei + xzhushi_fei + banla + gongyi + gongfei)

        invoice_s = str(invoice or '').strip()
        normalized = _normalize_order(invoice_s)

        if no == '19楼' or invoice_s == '19楼':
            cat, fly_key = '部门-真诚', None
        elif not invoice_s:
            cat, fly_key = '现货', None
        elif ORDER_LIKE.match(invoice_s):
            cat, fly_key = '客户单', f'E-{normalized}'
        else:
            cat, fly_key = '客户单', invoice_s

        items.append({'no': no, '类别': cat, '单号': invoice, '品名': pinming,
                      '成色': material, '件数': qty,
                      '总重': total_weight,
                      '主石重量': main_w if main_w > 0 else None,
                      '副石重量': side_w if side_w > 0 else None,
                      '镶嵌成本': cost,
                      '飞书匹配键': fly_key})
    return items


PARSERS = {'A': parse_A, 'B': parse_B, 'D': parse_D, 'E': parse_E}


def get_parser(code):
    return PARSERS[code]
