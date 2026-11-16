"""开工前核对：这份文件的表头是不是这家工厂的出货单。

文件名只能猜工厂（例如含「郑国远」就当雅希），内容不对时解析器会把别的表格
按固定列位硬读，算出一堆乱数再写进完成单和流水（2026-11-16：一份银饰「出货明细表」
被当成雅希写进了广州流水）。所以解析前先看表头，对不上就整份拒绝，不猜。

每家只核对解析器真正依赖的那张表、那几行表头：
- 固定列位的解析器（雅希、二厂、黛宝天然、布心）：表头词必须在对应的列上
- 按表头找列的解析器（黛宝培育、猛哥）：表头词在第 4 行出现即可
允许少数几个词对不上（工厂偶尔改字），但大部分要对上。
"""
import re
import openpyxl

LABELS = {'A': '雅希', 'B': '二厂', 'D': '黛宝', 'E': '猛哥', '黛宝天然': '黛宝', '布心': '布心'}

# kind → (选表, 表头行范围, [(列号或 None, 表头词)], 至少对上几个)
SIGNATURES = {
    'A': ('first', (1, 5), [(3, '下单编号'), (4, '单号'), (5, '证书编号'), (6, '款式'), (7, '成色'),
                            (8, '数量'), (13, '折铂金'), (14, '折足金')], 6),
    'B': ('last', (1, 8), [(1, '序号'), (2, '品名'), (3, '单号'), (5, '成色'), (6, '件数'),
                           (7, '出货重'), (12, '主石'), (28, '合计')], 6),
    'D': ('active', (4, 4), [(None, w) for w in ('序号', '条码号', '款号', '件数', '名称', '成色',
                                                 '折足', '主石', '总金额')], 7),
    'E': ('last', (4, 4), [(None, w) for w in ('序号', '单号', '品名', '成色', '件数', '总重', '含耗金重',
                                               '金价', '金值', '主石', '副石', '总金额')], 8),
    '黛宝天然': ('出货单-', (4, 5), [(1, '序号'), (2, '款号'), (3, '条码号'), (5, '证书号'), (8, '名称'),
                                 (10, '件数'), (11, '总重'), (18, '折足'), (93, '总价')], 7),
    '布心': ('布心', (1, 12), [(2, '客户'), (3, '款号'), (4, '证书号'), (6, '名称'), (7, '件数'),
                           (8, '成色')], 5),
}
MAX_COL = 100


def _norm(value):
    return re.sub(r'\s+', '', str(value)) if value is not None else ''


def _sheets(wb, pick):
    names = wb.sheetnames
    if pick == 'first':
        return [wb[names[0]]]
    if pick == 'last':
        return [wb[names[-1]]]
    if pick == 'active':
        return [wb.active or wb[names[0]]]
    if pick == '出货单-':
        return [wb[next((n for n in names if n.startswith('出货单-')), names[0])]]
    return [wb[n] for n in names if 'PT出货单' in n or '18K出货单' in n]  # 布心


def _hits(ws, rows, anchors):
    by_col, everything = {}, []
    for row in ws.iter_rows(min_row=rows[0], max_row=rows[1], max_col=MAX_COL):
        for idx, cell in enumerate(row, 1):
            text = _norm(cell.value)
            if text:
                by_col.setdefault(idx, []).append(text)
                everything.append(text)
    return sum(any(word in text for text in (everything if col is None else by_col.get(col, ())))
               for col, word in anchors)


def _matches(wb, kind):
    pick, rows, anchors, need = SIGNATURES[kind]
    try:
        return any(_hits(ws, rows, anchors) >= need for ws in _sheets(wb, pick))
    except Exception:
        return False


def kind_for(code, is_natural=False, natural_factory=None):
    """页面上的工厂代号 → 要核对的模板。二厂真诚、猛哥真诚和培育单是同一种表。"""
    if code == 'BUXIN' or natural_factory == '布心':
        return '布心'
    if is_natural and code == 'D':
        return '黛宝天然'
    return code


def problem(xlsx_path, kind):
    """格式对得上返回 None；对不上返回一句给用户看的话。"""
    label = LABELS.get(kind, kind)
    if kind not in SIGNATURES:
        return None
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    except Exception:
        return f"这份文件打不开，不是{label}的出货单格式，没有处理"
    try:
        if _matches(wb, kind):
            return None
        others = sorted({LABELS[k] for k in SIGNATURES if LABELS[k] != label and _matches(wb, k)})
    finally:
        wb.close()
    hint = f"（内容像{others[0]}的单子，请核对文件名或手动选工厂）" if len(others) == 1 else ""
    return f"这份单子不是{label}的格式，没有处理{hint}"
