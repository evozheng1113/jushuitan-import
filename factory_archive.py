"""将本次完成单的数据快照追加到指定飞书电子表格；不写业务成本表。"""
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from zoneinfo import ZoneInfo
from io import BytesIO
import hashlib
import math
import json
import re
import time
import openpyxl
from openpyxl.utils.datetime import to_excel
import requests

TOKEN = "VWlWsPmI9h5YMgtmHlNcHdgHnFh"
BASE = "https://open.feishu.cn/open-apis"
NAMES = {"A": "雅希", "B": "二厂", "D": "黛宝", "E": "猛哥", "BUXIN": "布心"}
ALIASES = {"雅希": {"雅希", "广州", "广州厂", "雅希工厂"},
           "二厂": {"二厂", "二厂工厂", "倾城", "倾诚"},
           "黛宝": {"黛宝", "黛宝工厂"}, "猛哥": {"猛哥", "猛哥工厂"},
           "布心": {"布心", "布心工厂"}}
# 飞书这些接口上限 100 次/秒。格式请求可重发，并行发更快；失败的再逐个补发。
WORKERS = 16
GAP_ROWS = 2  # 每批之间空两行，和手工贴的 2026 工厂单一样
_SESSION = requests.Session()


class _Http:
    """复用连接（服务器在海外，每次重新握手很慢）。测试替换 requests.get/post/put 时仍走替换后的函数。"""
    def __getattr__(self, name):
        fn = getattr(requests, name)
        return getattr(_SESSION, name) if getattr(fn, "__module__", "").startswith("requests") else fn


HTTP = _Http()
# 内容中断这么多行以后的零散单元格（工厂误输入）不算本批。
STRAY_GAP = 20
# 这两家一个文件有很多sheet，最后一个才是新出货；其余是历史，不能再次归档。
LATEST_SHEET_ONLY = {"二厂", "猛哥"}
# 这些工厂的货品图放进单元格。布心等的图片不是货品图，不复制。
IMAGE_FACTORIES = {"雅希"}


def date_text(value, fmt):
    """日期按工厂单显示的样子写入流水；单号 11-2-3 被 Excel 变成日期时也能还原。"""
    pattern = (fmt or "").replace("\\", "").lower()
    if pattern == "mm-dd-yy":  # Excel 内置短日期
        return f"{value.year}/{value.month}/{value.day}"
    if not re.fullmatch(r"(yyyy|yy|m{1,2}|d{1,2})([-/.年月日]+(yyyy|yy|m{1,2}|d{1,2}))*[年月日]?", pattern):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    parts = {"yyyy": str(value.year), "yy": f"{value.year % 100:02d}", "mm": f"{value.month:02d}",
             "m": str(value.month), "dd": f"{value.day:02d}", "d": str(value.day)}
    return re.sub(r"yyyy|yy|mm|m|dd|d", lambda m: parts[m[0]], pattern)


def scalar(value, percent=False):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("完成文件包含无效数字")
    if percent and isinstance(value, (int, float)):
        return f"{value * 100:.2f}%"
    if not isinstance(value, (str, int, float, bool)):
        return str(value)  # 时间等类型转成文字，避免整批写入失败
    # 原表公式只复制计算结果；文本不得变成目标表里的公式。
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


class CompletedSheetRows(list):
    """接口保持兼容：值就是完成单；格式只保存在内存，不增加可见列。"""
    def __init__(self):
        super().__init__()
        self.styles = []
        self.merges = []
        self.heights = []
        self.widths = {}
        self.batch = ""  # 本批原始内容的指纹，用来判断同一批是否写过
        self.images = []  # (本批内0起的行, 列号, 图片内容, 文件名)
        self.blocks = []  # (sheet名, 起始行, 结束行, 列数, 合并范围, 在本批中的起始位置)
        self.image_failed = 0
        self.old_dates = {}  # 旧版把日期写成完整时间，修复格式时仍要认得出


def rgb(color):
    if color is not None and color.type == "rgb" and isinstance(color.rgb, str):
        return "#" + color.rgb[-6:]
    if color is not None and color.type == "indexed":
        from openpyxl.styles.colors import COLOR_INDEX
        if 0 <= color.indexed < len(COLOR_INDEX):
            return "#" + COLOR_INDEX[color.indexed][-6:]
    return None


def compatible_formatter(fmt):
    """Only send simple numeric patterns, never Excel locale/accounting syntax."""
    if not fmt or fmt == "General":
        return None
    section = fmt.split(";")[0]
    section = re.sub(r"\[[^\]]*\]", "", section)
    if re.search(r"[ymdhs]", section, re.I):
        return None  # dates are already ISO strings
    if "%" in section:
        match = re.search(r"0\.([0#]+)", section)
        return "0" + ("." + "0" * min(len(match[1]), 8) if match else "") + "%"
    match = re.fullmatch(r"(?:#,##)?0(?:\.([0#]{1,8}))?", section)
    if match:
        return section
    return None


class ArchiveAPIError(RuntimeError):
    def __init__(self, code, message):
        self.code = code
        super().__init__(f"飞书归档失败：{code} {message}")


class ArchiveUncertain(RuntimeError):
    """写入请求已发出但结果未确认，可能已写入，不能当作失败重跑。"""


class LayoutIncomplete(RuntimeError):
    def __init__(self, title, sid, first_row, rows, cause):
        self.sid, self.first_row, self.rows = sid, first_row, rows
        super().__init__(f"流水已追加到{title}第{first_row}行起，格式未完成：{cause}；请使用仅修复归档格式，请勿重跑成本")


def font_size(size):
    """飞书只接受 9~36 的整数字号；原表更小或更大时取最接近的值。"""
    return f"{min(36, max(9, round(size)))}pt/1.5"


def batch_end(rows, needed=0):
    """rows 为有实际内容的行号。工厂解析到的货品行（needed）绝不截断。"""
    end = 0
    for row in sorted(rows):
        if end and row - end > STRAY_GAP and row > needed:
            break
        end = row
    return max(end, needed)


def item_last_row(group, rows):
    """货品占用、且确实有内容的最后一行；解析器可能把后面的空行也算给最后一件。"""
    return max([it["row"] for it in group if it.get("row") in rows]
               + [r for it in group for r in (it.get("rows") or []) if r in rows] + [0])


def default_repair_sheets(sheetnames, factory):
    """修复格式时默认选中的sheet要和当初归档的范围一致。"""
    return list(sheetnames[-1:]) if NAMES.get(factory, factory) in LATEST_SHEET_ONLY else list(sheetnames)


def cell_style(cell):
    style = {"font": {"bold": bool(cell.font.bold), "italic": bool(cell.font.italic)}}
    if cell.font.sz:
        style["font"]["fontSize"] = font_size(cell.font.sz)
    foreground = rgb(cell.font.color)
    background = rgb(cell.fill.fgColor) if cell.fill.patternType == "solid" else None
    if foreground:
        style["foreColor"] = foreground
    if background:
        style["backColor"] = background
    formatter = compatible_formatter(cell.number_format)
    if formatter:
        style["formatter"] = formatter
    if cell.alignment.horizontal in ("left", "center", "right"):
        style["hAlign"] = {"left": 0, "center": 1, "right": 2}[cell.alignment.horizontal]
    if cell.alignment.vertical in ("top", "center", "bottom"):
        style["vAlign"] = {"top": 0, "center": 1, "bottom": 2}[cell.alignment.vertical]
    sides = [cell.border.left, cell.border.right, cell.border.top, cell.border.bottom]
    if all(side is not None and side.style for side in sides):
        style["borderType"] = "FULL_BORDER"
        style["borderColor"] = rgb(sides[0].color) or "#000000"
    return style


def sheet_images(ws, offset, start, end, width):
    """取本批范围内的图片。图片要落在合并格的左上角，同一格只放第一张。"""
    found, used = [], set()
    merged = {(r, c): (m.min_row, m.min_col) for m in ws.merged_cells.ranges
              for r in range(m.min_row, m.max_row+1) for c in range(m.min_col, m.max_col+1)}
    for image in ws._images:
        anchor = getattr(image.anchor, "_from", None)
        if anchor is None:
            continue
        row, col = merged.get((anchor.row+1, anchor.col+1), (anchor.row+1, anchor.col+1))
        if not start <= row <= end or col > width or (row, col) in used or ws.cell(row, col).value is not None:
            continue  # 写图片会盖掉格子内容，有内容的格子不写
        used.add((row, col))
        found.append((offset+row-start, col, image._data(), f"{row}_{col}.{image.format or 'png'}"))
    return found


def write_images(client, sid, rows, first_row, tick=None):
    """图片失败只记数，数字文字已写好，这里绝不抛错。"""
    headers = client._headers(json_body=True)

    def task(index, col, data, name):
        cell = f"{openpyxl.utils.get_column_letter(col)}{first_row+index}"
        return lambda: layout_call("post", f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/values_image", headers=headers,
                                   json={"range": f"{sid}!{cell}:{cell}", "image": list(data), "name": name}, timeout=60)

    failed = 0
    for retry in run_parallel([task(*image) for image in rows.images], tick):
        try:
            retry()
        except (ArchiveAPIError, requests.exceptions.RequestException, ValueError):
            failed += 1
    rows.image_failed = failed


def snapshot(data, source_path, items, factory, filename, flow):
    """复制本次完成单的原表块，不添加任何归档字段，不挤掉空列。"""
    books = []
    try:
        for source, cached in ((BytesIO(data), True), (BytesIO(data), False),
                               (source_path, True), (source_path, False)):
            books.append(openpyxl.load_workbook(source, data_only=cached))
        wb, formulas, original, original_formulas = books
        grouped = {}
        for item in items:
            sn = item.get("_sheet") or item.get("_sheet_title")
            if NAMES.get(factory, factory) in LATEST_SHEET_ONLY:
                sn = wb.sheetnames[-1]
            elif not sn and NAMES.get(factory, factory) == "黛宝":
                sn = wb.active.title  # 与黛宝解析器一致
            elif not sn:
                sn = wb.sheetnames[-1] if "row" in item else wb.sheetnames[0]
            if sn not in wb.sheetnames:
                raise ValueError(f"完成单找不到原sheet：{sn}")
            grouped.setdefault(sn, []).append(item)
        result = CompletedSheetRows()
        batch = hashlib.sha256()
        for sn, group in grouped.items():
            ws, fs = wb[sn], formulas[sn]
            populated = [(r, c) for (r, c), cell in fs._cells.items()
                         if cell.value is not None and str(cell.value).strip()]
            if not populated:
                continue
            content_rows = {r for r, c in populated}
            start, end = 1, batch_end(content_rows, item_last_row(group, content_rows))
            if NAMES.get(factory, factory) == "布心":
                # 与布心解析器一致，最后一个“客户”表头至最后一个“合计”为本次批次。
                totals = [r for r, c in populated if c == 2 and "合计" in str(fs.cell(r, c).value)]
                if not totals:
                    raise ValueError(f"{sn}没有本批合计行，未归档")
                last_total = max(totals)
                headers = [r for r, c in populated if c == 2 and r < last_total and str(fs.cell(r, c).value).strip() == "客户"]
                if not headers:
                    raise ValueError(f"{sn}没有本批表头，未归档")
                header = max(headers)
                dates = [r for r, c in populated if c == 2 and r < header and re.search(r"20\d{2}年\d{1,2}月\d{1,2}日", str(fs.cell(r, c).value))]
                start = max(dates) if dates else header
                if any(not (header < it["row"] < last_total) for it in group):
                    raise ValueError(f"{sn}本次行号与批次边界不符，未归档")
            relevant = [(r, c) for r, c in populated if start <= r <= end]
            width = max(c for r, c in relevant)
            # 保留与内容相交的合并表头/表尾，排除远处仅带格式的空白。
            merges = [m for m in fs.merged_cells.ranges if m.min_row <= end and m.max_row >= start
                      and fs.cell(m.min_row, m.min_col).value is not None]
            for merged in merges:
                if merged.min_row < start:
                    raise ValueError(f"{sn}的合并单元格跨批次，未归档")
                width = max(width, merged.max_col)
                end = max(end, merged.max_row)
            if width > 500 or end-start+1 > 4998:
                raise ValueError("完成单范围过大，未归档，请核对原文件的有效范围")
            # 用原出货单计算，金价或成本变化不影响判断；行号取相对值，不受批次位置影响。
            source = original[sn] if sn in original.sheetnames else ws
            for (r, c), cell in sorted(source._cells.items()):
                if start <= r <= end and cell.value is not None:
                    batch.update(json.dumps([r-start, c, str(cell.value)], ensure_ascii=False).encode())
            offset = len(result)
            if NAMES.get(factory, factory) in IMAGE_FACTORIES:
                result.images.extend(sheet_images(source, offset, start, end, width))
            for r in range(start, end + 1):
                values = []
                row_styles = []
                for c in range(1, width+1):
                    cell = ws.cell(r, c)
                    formula = fs.cell(r, c)
                    value = cell.value
                    if formula.data_type == "f" and value is None:
                        if sn in original.sheetnames and original_formulas[sn].cell(r, c).value == formula.value:
                            value = original[sn].cell(r, c).value
                        if value is None:
                            raise ValueError(f"{sn}!{formula.coordinate}公式没有计算结果，无法原样归档；请先用Excel重算保存")
                    if isinstance(value, (datetime, date)):
                        shown = date_text(value, formula.number_format)
                        result.old_dates[(offset+r-start, c-1)] = (
                            scalar(value), to_excel(value))
                        values.append(shown)
                    else:
                        values.append(scalar(value))
                    row_styles.append(cell_style(formula))
                result.append(values)
                # 将同一行相邻的相同样式压成一个范围。
                left = 1
                for c in range(2, width+2):
                    if c == width+1 or row_styles[c-1] != row_styles[left-1]:
                        result.styles.append((offset+r-start+1, left, c-1, row_styles[left-1]))
                        left = c
                height = fs.row_dimensions[r].height or fs.sheet_format.defaultRowHeight or 15
                result.heights.append((offset+r-start+1, max(1, round(height * 4/3))))
            for merged in merges:
                result.merges.append((offset+merged.min_row-start+1, merged.min_col,
                                      offset+merged.max_row-start+1, merged.max_col))
            result.blocks.append((sn, start, end, width,
                                  [(m.min_row, m.min_col, m.max_row, m.max_col) for m in merges], offset))
            for c in range(1, width+1):
                # 同工厂同一sheet的列宽为共享属性，记录完成单使用的宽度。
                dim = next((d for d in fs.column_dimensions.values() if (d.min or 0) <= c <= (d.max or 0)), None)
                chars = dim.width if dim is not None else fs.sheet_format.defaultColWidth or 8.43
                result.widths[c] = max(1, round(chars * 7 + 5))
            result.append([""] * width)  # 两批之间只留一空行，不增加说明文字。
        if not result:
            raise ValueError("本次没有可归档的完成单")
        if len(result) > 5000:
            raise ValueError("本次完成单超过5000行，未归档")
        width = max(map(len, result))
        for row in result:
            row.extend([""] * (width-len(row)))
        result.batch = batch.hexdigest()[:32]
        return result
    finally:
        for book in books:
            book.close()


def rejected(exc):
    """飞书表格接口拒绝了某项内容（90xxx）。登录、权限类错误不算，要继续报错。"""
    return str(exc.code).startswith("90") and len(str(exc.code)) == 5


def layout_call(method, url, **kwargs):
    """格式请求可安全重发；飞书限速时等一下再试，不追加、不改写数据。"""
    for attempt in range(5):
        response = getattr(HTTP, method)(url, **kwargs)
        if getattr(response, "status_code", 200) == 429 and attempt < 4:
            time.sleep(2 * (attempt+1))
            continue
        try:
            return checked(response)
        except ArchiveAPIError as exc:
            if str(exc.code) != "99991400" or attempt == 4:
                raise
            time.sleep(2 * (attempt+1))


def run_parallel(tasks, tick=None):
    """tasks 里每项可安全重发。返回并行失败、需要逐个补发的任务。"""
    failed = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(task): task for task in tasks}
        for future in as_completed(futures):
            try:
                future.result()
            except Exception:
                failed.append(futures[future])
            if tick:
                tick()  # 在主线程更新进度，页面组件不能从子线程调用
    return failed


def same_value(got, want):
    if isinstance(got, (list, dict)) and want in ("", None):
        return True  # 空格子里读到对象，是写进去的图片
    if isinstance(got, (int, float)) and isinstance(want, (int, float)):
        return abs(got - want) < 1e-6
    return str("" if got is None else got) == str("" if want is None else want)


def checked(response):
    response.raise_for_status()
    payload = response.json()
    if payload.get("code") != 0:
        raise ArchiveAPIError(payload.get("code"), payload.get("msg", ""))
    return payload.get("data") or {}


def append_archive(client, rows, factory, progress=None):
    if client.spreadsheet_token != TOKEN:
        raise ValueError("归档目标表不正确")
    factory = NAMES.get(factory, factory)
    listing = checked(HTTP.get(f"{BASE}/sheets/v3/spreadsheets/{TOKEN}/sheets/query",
                                  headers=client._headers(), timeout=30))
    sheets = listing.get("sheets", [])
    accepted = ALIASES.get(factory, {factory})
    matches = [s for s in sheets if s.get("title", "").strip() in accepted]
    if len(matches) != 1:
        titles = "、".join(s.get("title", "") for s in sheets)
        raise RuntimeError(f"无法唯一确定{factory}的sheet。现有sheet：{titles}；请将对应sheet命名为{factory}")
    sheet = matches[0]
    sid = sheet["sheet_id"]
    width = max(map(len, rows))
    grid = sheet.get("grid_properties") or {}
    current_columns = grid.get("column_count")
    if not isinstance(current_columns, int) or current_columns < 1:
        raise RuntimeError(f"未能读取{sheet['title']}的列数，未追加数据")
    if current_columns < width:
        # 飞书官方增加行列接口：只在右侧添加空列，不改动已有数据。
        extra = width - current_columns
        problem = None
        try:
            expansion = checked(HTTP.post(
                f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/dimension_range",
                headers=client._headers(json_body=True),
                json={"dimension": {"sheetId": sid, "majorDimension": "COLUMNS", "length": extra}},
                timeout=30))
            if expansion.get("addCount") != extra:
                problem = "补列返回数量未确认"
        except requests.exceptions.RequestException as exc:
            response = getattr(exc, "response", None)
            status = getattr(response, "status_code", None)
            problem = f"补列请求异常：{type(exc).__name__}" + (f"，HTTP {status}" if status else "")
            if response is not None:
                try:
                    error = response.json()
                    problem += f"，code={error.get('code')}，{str(error.get('msg', ''))[:200]}"
                except (ValueError, AttributeError):
                    pass
        except RuntimeError as exc:
            problem = str(exc)
        if problem:
            # 补列的请求可能已经生效。仅重新读取，不重发补列或流水写入。
            try:
                latest = checked(HTTP.get(
                    f"{BASE}/sheets/v3/spreadsheets/{TOKEN}/sheets/query",
                    headers=client._headers(), timeout=30))
                found = next((s for s in latest.get("sheets", []) if s.get("sheet_id") == sid), None)
                latest_grid = (found or {}).get("grid_properties") or {}
                actual = latest_grid.get("column_count")
            except (requests.exceptions.RequestException, RuntimeError):
                raise RuntimeError(f"{problem}；重新读取列数也失败，本次尚未追加流水")
            if not isinstance(actual, int) or actual < width:
                raise RuntimeError(f"{problem}；{sheet['title']}目前{actual}列，需要{width}列，本次尚未追加流水")
            grid = latest_grid
    total_rows = max(int(grid.get("row_count", 1)), 1)
    columns = max(int(grid.get("column_count") or width), width)
    used = last_used_row(client, sid, total_rows, columns)
    first_row = used + GAP_ROWS + 1 if used else 1
    need = first_row + len(rows) - 1
    while total_rows < need:
        # 只在表格底部加空行，不动已有数据；这一步失败说明还没写任何流水。
        add = min(need - total_rows, 5000)
        checked(HTTP.post(f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/dimension_range",
                          headers=client._headers(json_body=True),
                          json={"dimension": {"sheetId": sid, "majorDimension": "ROWS", "length": add}}, timeout=30))
        total_rows += add
    # 写到指定行，不用飞书“找第一个空格”的追加方式（会插进旧数据中间）。每次不超过 5000 行、100 列。
    wrote = False
    try:
        for top in range(0, len(rows), 5000):
            block = rows[top:top + 5000]
            for left in range(0, width, 100):
                right = min(left + 100, width)
                target = (f"{sid}!{openpyxl.utils.get_column_letter(left + 1)}{first_row + top}:"
                          f"{openpyxl.utils.get_column_letter(right)}{first_row + top + len(block) - 1}")
                checked(HTTP.put(f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/values",
                                 headers=client._headers(json_body=True),
                                 json={"valueRange": {"range": target, "values": [r[left:right] for r in block]}},
                                 timeout=60))
                wrote = True
    except requests.exceptions.RequestException as exc:
        raise ArchiveUncertain("写流水的请求没有收到确认，可能已写入；请先查看工厂流水再决定是否重跑") from exc
    except ArchiveAPIError as exc:
        if wrote:
            raise ArchiveUncertain(f"流水只写入了一部分（{exc}）；请先查看工厂流水") from exc
        raise
    rows.first_row = first_row
    if used:
        rows.widths = {}  # 列宽对整张表生效，表里已有内容时沿用原列宽
    if isinstance(rows, CompletedSheetRows):
        layout_error = None
        tick = progress_ticker(rows, progress)
        try:
            apply_layout(client, sid, rows, first_row, tick)
        except Exception as exc:
            layout_error = exc
        write_images(client, sid, rows, first_row, tick)
        if layout_error is not None:
            raise LayoutIncomplete(sheet["title"], sid, first_row, rows, layout_error) from layout_error
    return sheet["title"]


def last_used_row(client, sid, total_rows, columns):
    """从表格底部往上读，找到真正有内容的最后一行（任何一列有值都算，包括图片）。"""
    right = openpyxl.utils.get_column_letter(min(max(columns, 1), 200))
    end = total_rows
    while end >= 1:
        start = max(1, end - 1999)
        data = checked(HTTP.get(f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/values/{sid}!A{start}:{right}{end}",
                                headers=client._headers(), params={"valueRenderOption": "UnformattedValue"},
                                timeout=60))
        values = (data.get("valueRange") or {}).get("values") or []
        for offset in range(len(values) - 1, -1, -1):
            if any(v not in (None, "") for v in (values[offset] or [])):
                return start + offset
        end = start - 1
    return 0


def progress_ticker(rows, progress):
    if not progress:
        return None
    total = len(rows.merges) + len(rows.images) + len(rows.heights) + len(rows.widths) + 1
    done = [0]

    def tick():
        done[0] += 1
        progress(min(done[0], total), total)
    return tick


def apply_layout(client, sid, rows, first_row, tick=None):
    def address(top, left, bottom, right):
        return f"{sid}!{openpyxl.utils.get_column_letter(left)}{first_row+top-1}:{openpyxl.utils.get_column_letter(right)}{first_row+bottom-1}"
    base = f"{BASE}/sheets/v2/spreadsheets/{TOKEN}"
    groups = {}
    for row, left, right, style in rows.styles:
        key = json.dumps(style, sort_keys=True)
        group = groups.setdefault(key, {"ranges": [], "style": style})
        for c in range(left, right+1, 100):
            group["ranges"].append(address(row, c, row, min(right, c+99)))
    batches = []
    for group in groups.values():
        for i in range(0, len(group["ranges"]), 100):
            batches.append({"ranges": group["ranges"][i:i+100], "style": group["style"]})
    def put_styles(data):
        layout_call("put", base + "/styles_batch_update", headers=client._headers(json_body=True),
                    json={"data": data}, timeout=60)

    def downgrade(group):
        # 每次只去掉一项可能被拒的样式；只设置样式，不追加、不改写业务数据。
        style = dict(group["style"])
        steps = [lambda s: s.pop("formatter", None),
                 lambda s: s.get("font", {}).pop("fontSize", None),
                 lambda s: (s.pop("foreColor", None), s.pop("backColor", None)),
                 lambda s: (s.pop("borderType", None), s.pop("borderColor", None)),
                 lambda s: (s.pop("hAlign", None), s.pop("vAlign", None))]
        style["font"] = dict(style.get("font", {}))
        for step in steps:
            step(style)
            try:
                put_styles([{"ranges": group["ranges"], "style": style}])
                return
            except ArchiveAPIError as exc:
                if not rejected(exc):
                    raise

    for i in range(0, len(batches), 20):
        batch = batches[i:i+20]
        try:
            put_styles(batch)
        except ArchiveAPIError as exc:
            if not rejected(exc):
                raise
            rows.format_fallback = True
            for group in batch:
                try:
                    put_styles([group])
                except ArchiveAPIError as single:
                    if not rejected(single):
                        raise
                    downgrade(group)
    headers = client._headers(json_body=True)

    def finish(tasks):
        # 并行失败的逐个补发；飞书明确拒绝的跳过，登录权限问题继续报错。
        for retry in run_parallel(tasks, tick):
            try:
                retry()
            except ArchiveAPIError as exc:
                if not rejected(exc):
                    raise
                rows.format_fallback = True

    def merge(top, left, bottom, right):
        return lambda: layout_call("post", base + "/merge_cells", headers=headers, timeout=30,
                                   json={"range": address(top, left, bottom, right), "mergeType": "MERGE_ALL"})

    def resize(major, start, end, pixels):
        return lambda: layout_call("put", base + "/dimension_range", headers=headers, timeout=30,
                                   json={"dimension": {"sheetId": sid, "majorDimension": major,
                                                       "startIndex": start-1, "endIndex": end},
                                         "dimensionProperties": {"fixedSize": pixels}})

    finish([merge(*item) for item in rows.merges])
    # 行高只设置新追加部分；列宽是飞书sheet全局属性。
    tasks = []
    for major, dimensions in (("ROWS", rows.heights), ("COLUMNS", sorted(rows.widths.items()))):
        groups = []
        for index, pixels in dimensions:
            index = first_row + index - 1 if major == "ROWS" else index
            if groups and groups[-1][1] == index-1 and groups[-1][2] == pixels:
                groups[-1][1] = index
            else:
                groups.append([index, index, pixels])
        tasks += [resize(major, start, end, pixels) for start, end, pixels in groups]
    finish(tasks)


def repair_layout(client, rows, factory, first_row):
    """Read/compare the full existing block before any layout mutation. No append."""
    if client.spreadsheet_token != TOKEN or not isinstance(first_row, int) or first_row < 1:
        raise ValueError("修复目标不正确")
    listing = checked(requests.get(f"{BASE}/sheets/v3/spreadsheets/{TOKEN}/sheets/query", headers=client._headers(), timeout=30))
    name = NAMES.get(factory, factory)
    matches = [s for s in listing.get("sheets", []) if s.get("title", "").strip() in ALIASES.get(name, {name})]
    if len(matches) != 1:
        raise ValueError("无法唯一确定修复的工厂sheet")
    sid = matches[0]["sheet_id"]
    width = max(map(len, rows))
    # Read unformatted values to compare numbers rather than currency display strings.
    for start in range(0, len(rows), 200):
        block = rows[start:start+200]
        for left in range(0, width, 100):
            right = min(left+100, width)
            area = f"{sid}!{openpyxl.utils.get_column_letter(left+1)}{first_row+start}:{openpyxl.utils.get_column_letter(right)}{first_row+start+len(block)-1}"
            data = checked(requests.get(f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/values/{area}", headers=client._headers(), params={"valueRenderOption":"UnformattedValue"}, timeout=30))
            actual = data.get("valueRange", {}).get("values")
            if actual is None:
                raise ValueError("无法核对已有流水，未修改格式")
            for r, expected in enumerate(block):
                line = actual[r] if r < len(actual) else []
                for c in range(left, right):
                    got = line[c-left] if c-left < len(line) else ""
                    want = expected[c]
                    legacy = getattr(rows, "old_dates", {}).get((start+r, c), ())
                    if not same_value(got, want) and not any(same_value(got, old) for old in legacy):
                        raise ValueError(f"第{first_row+start+r}行第{c+1}列与完成文件不一致，未修改；请核对文件、sheet和起始行")
    apply_layout(client, sid, rows, first_row)
    write_images(client, sid, rows, first_row)  # 中断过的批次在这里补图，同格重写不会重复
    return matches[0]["title"]
