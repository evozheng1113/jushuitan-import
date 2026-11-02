"""将本次完成单的数据快照追加到指定飞书电子表格；不写业务成本表。"""
from datetime import date, datetime
from io import BytesIO
import math
import openpyxl
import requests

TOKEN = "VWlWsPmI9h5YMgtmHlNcHdgHnFh"
BASE = "https://open.feishu.cn/open-apis"
NAMES = {"A": "雅希", "B": "二厂", "D": "黛宝", "E": "猛哥", "BUXIN": "布心"}
ALIASES = {"雅希": {"雅希", "广州厂", "雅希工厂"},
           "二厂": {"二厂", "二厂工厂", "倾城", "倾诚"},
           "黛宝": {"黛宝", "黛宝工厂"}, "猛哥": {"猛哥", "猛哥工厂"},
           "布心": {"布心", "布心工厂"}}
HEAD = ["运行时间", "完成文件", "流程", "工厂", "原sheet", "原行号", "下单单号", "证书编号",
        "总镶嵌成本", "客户名称", "利润", "利润率"]


def scalar(value, percent=False):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("完成文件包含无效数字")
    if percent and isinstance(value, (int, float)):
        return f"{value * 100:.2f}%"
    # 原表公式只复制计算结果；文本不得变成目标表里的公式。
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def snapshot(data, source_path, items, factory, filename, flow):
    """每次只归档本次解析的行，含完成单原列和明确标注的成本/利润。"""
    factory = NAMES.get(factory, factory)
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    formulas = openpyxl.load_workbook(BytesIO(data), data_only=False)
    original = openpyxl.load_workbook(source_path, data_only=True)
    original_formulas = openpyxl.load_workbook(source_path, data_only=False)
    records = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    seen = set()
    try:
        for it in items:
            sn = it.get("_sheet") or it.get("_sheet_title")
            if not sn:
                sn = wb.sheetnames[-1] if "row" in it else wb.sheetnames[0]
            if sn not in wb.sheetnames:
                raise ValueError(f"完成单找不到原sheet：{sn}")
            rows = [it["row"]] if "row" in it else it.get("rows", [])
            if not rows:
                raise ValueError("出货记录缺少原行号，未归档")
            ws, fs = wb[sn], formulas[sn]
            for index, r in enumerate(rows):
                if (sn, r) in seen:
                    raise ValueError(f"本次记录重复引用原行：{sn} 第{r}行")
                seen.add((sn, r))
                # 带格式的空列不计入宽度。
                populated = [c for (rr, c), cell in fs._cells.items() if rr == r and cell.value is not None]
                width = max(populated, default=1)
                if width > 500:
                    raise ValueError("原单有效列超过500列，未归档")
                raw = []
                for c in range(1, width + 1):
                    cell = ws.cell(r, c)
                    value = cell.value
                    formula = fs.cell(r, c)
                    if formula.data_type == "f" and value is None:
                        if sn in original.sheetnames and original_formulas[sn].cell(r, c).value == formula.value:
                            value = original[sn].cell(r, c).value
                        if value is None:
                            # 不伪装成零，也不把移位公式写入飞书。
                            value = "[原表公式无缓存]"
                    raw.append(scalar(value, "%" in cell.number_format))
                first = index == 0
                order = it.get("下单单号") or it.get("飞书匹配键") or it.get("下单编号") or ""
                customer = it.get("飞书客户名") or (it.get("款号") if "row" in it else "") or ""
                records.append([now, filename, flow, factory, sn, r, scalar(order), scalar(it.get("证书编号") or it.get("证书号")),
                                scalar(it.get("镶嵌成本")) if first else "",
                                scalar(customer), scalar(it.get("飞书利润")) if first else "",
                                scalar(it.get("飞书利润率"), True) if first else ""] + raw)
        if not records:
            raise ValueError("本次没有可归档的出货记录")
        full_width = max(map(len, records))
        used = list(range(len(HEAD))) + [c for c in range(len(HEAD), full_width)
                if any(c < len(r) and r[c] != "" for r in records)]
        header = HEAD + [f"原表{openpyxl.utils.get_column_letter(c-len(HEAD)+1)}列" for c in used[len(HEAD):]]
        result = [header] + [[r[c] if c < len(r) else "" for c in used] for r in records]
        width = len(header)
        if len(result) > 5000 or width > 100:
            raise ValueError("本次归档超过单次5000行或100列限制，未写入")
        return result
    finally:
        for book in (wb, formulas, original, original_formulas):
            book.close()


def checked(response):
    response.raise_for_status()
    payload = response.json()
    if payload.get("code") != 0:
        raise RuntimeError(f"飞书归档失败：{payload.get('code')} {payload.get('msg', '')}")
    return payload.get("data") or {}


def append_archive(client, rows, factory):
    if client.spreadsheet_token != TOKEN:
        raise ValueError("归档目标表不正确")
    factory = NAMES.get(factory, factory)
    listing = checked(requests.get(f"{BASE}/sheets/v3/spreadsheets/{TOKEN}/sheets/query",
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
    if grid.get("column_count", 0) < width:
        raise RuntimeError(f"{sheet['title']}需要至少{width}列，请先扩展工作表列数")
    last = max(int(grid.get("row_count", 1)), 1)
    col = openpyxl.utils.get_column_letter(width)
    # 整个已分配范围搜索末尾，避免只搜索首行覆盖旧流水。单次请求，不自动重试写入。
    try:
        result = checked(requests.post(
            f"{BASE}/sheets/v2/spreadsheets/{TOKEN}/values_append?insertDataOption=INSERT_ROWS",
            headers=client._headers(json_body=True),
            json={"valueRange": {"range": f"{sid}!A1:{col}{last}", "values": rows}}, timeout=60))
    except requests.exceptions.RequestException as exc:
        raise RuntimeError("归档请求未获得确认，可能已写入；请先查看工厂流水再决定是否重跑") from exc
    updates = result.get("updates") or {}
    if updates.get("updatedRows") != len(rows):
        raise RuntimeError("飞书未确认完整归档，请检查工厂流水；勿直接重跑")
    return sheet["title"]
