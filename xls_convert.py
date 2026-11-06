"""把老 .xls 工厂单转成 .xlsx，保留内容和主要格式。线上没有 LibreOffice，只用纯 Python。"""
from io import BytesIO
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import xlrd

HORIZONTAL = {1: "left", 2: "center", 3: "right"}
VERTICAL = {0: "top", 1: "center", 2: "bottom"}


def _color(book, index):
    rgb = book.colour_map.get(index)
    return "FF%02X%02X%02X" % rgb if rgb else None


def _value(book, cell):
    kind, value = cell.ctype, cell.value
    if kind in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if kind == xlrd.XL_CELL_NUMBER:
        return int(value) if float(value).is_integer() else value
    if kind == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(value, book.datemode)
        except (ValueError, OverflowError):
            return value
    if kind == xlrd.XL_CELL_BOOLEAN:
        return bool(value)
    if kind == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(value, "#ERROR")
    return value


def _style(book, target, xf):
    font = book.font_list[xf.font_index]
    target.font = Font(sz=font.height / 20, b=bool(font.bold), i=bool(font.italic),
                       color=_color(book, font.colour_index))
    fmt = book.format_map.get(xf.format_key)
    if fmt and fmt.format_str:
        target.number_format = fmt.format_str
    target.alignment = Alignment(horizontal=HORIZONTAL.get(xf.alignment.hor_align),
                                 vertical=VERTICAL.get(xf.alignment.vert_align),
                                 wrap_text=bool(xf.alignment.text_wrapped))
    if xf.background.fill_pattern == 1:
        background = _color(book, xf.background.pattern_colour_index)
        if background:
            target.fill = PatternFill("solid", fgColor=background)
    line = xf.border
    side = lambda style: Side(style="thin") if style else Side()
    target.border = Border(left=side(line.left_line_style), right=side(line.right_line_style),
                           top=side(line.top_line_style), bottom=side(line.bottom_line_style))


def xls_to_xlsx(data):
    with open(os.devnull, "w") as quiet:
        book = xlrd.open_workbook(file_contents=data, formatting_info=True, logfile=quiet)
    out = openpyxl.Workbook()
    out.remove(out.active)
    for sheet in book.sheets():
        ws = out.create_sheet(sheet.name[:31])
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                source = sheet.cell(r, c)
                target = ws.cell(r + 1, c + 1, _value(book, source))
                if source.xf_index is not None:
                    _style(book, target, book.xf_list[source.xf_index])
        for top, bottom, left, right in sheet.merged_cells:
            ws.merge_cells(start_row=top + 1, end_row=bottom, start_column=left + 1, end_column=right)
        for c, info in sheet.colinfo_map.items():
            ws.column_dimensions[get_column_letter(c + 1)].width = info.width / 256
        for r, info in sheet.rowinfo_map.items():
            if info.height:
                ws.row_dimensions[r + 1].height = info.height / 20
    buffer = BytesIO()
    out.save(buffer)
    return buffer.getvalue()
