"""聚水潭入库 Excel 生成 — 网站版 v11
- 接飞书读: 客户名 / 利润 / 圈号 / 主石 / 裸钻成本 / 配石成本 / 证书编码
- 不写飞书 (网站默认只读, 避免误更)
- 凭证: Streamlit Cloud Secrets 加密存
"""
import streamlit as st
import tempfile, os, traceback, re, time
from datetime import datetime

import factories
import jushuitan_import as jst
import natural  # 天然钻 (v19)
from feishu_client import FeishuClient, APP_TOKEN, TABLE_ID, load_credentials

# 培育钻 code(A/B/D/E) → natural.py 里对应的天然钻工厂名
NATURAL_FACTORY_MAP = {
    'B': '二厂',   # 二厂真诚
    'D': '黛宝',
    'E': '猛哥',
    # A(雅希) 暂无天然钻流程 (待需求)
}
NATURAL_FACTORY_MAP_BUXIN = '布心'  # 布心走独立文件, 靠文件名单独识别


st.set_page_config(page_title="聚水潭入库生成", page_icon="💎", layout="centered")

st.title("💎 聚水潭入库 Excel 生成")
st.caption("上传工厂出货单 → 查飞书 → 生成聚水潭批量入库模板")

# ---------------- 飞书工具 ----------------
def _fetch_after_update(client, key, target_cost, max_wait=4):
    """v14.2: 写完镶嵌成本后, 轮询飞书等公式刷新, 返回 (record, 是否成功刷新)
    判断标准: 镶嵌成本字段已经变成 target_cost, 再等 0.5s 让利润公式算完
    """
    deadline = time.time() + max_wait
    last_rec = None
    while time.time() < deadline:
        try:
            rec = client.find_by_order_number(APP_TOKEN, TABLE_ID, key)
        except Exception:
            rec = None
        if rec:
            last_rec = rec
            current_cost = client.get_number(rec['fields'].get('镶嵌成本'))
            if current_cost is not None and abs(current_cost - target_cost) < 0.5:
                time.sleep(0.5)
                try:
                    rec3 = client.find_by_order_number(APP_TOKEN, TABLE_ID, key)
                except Exception:
                    rec3 = None
                return (rec3 or rec, True)
        time.sleep(0.5)
    return (last_rec, False)


# ---------------- 飞书初始化 ----------------
@st.cache_resource
def get_feishu_client():
    app_id, app_secret = load_credentials()
    return FeishuClient(app_id, app_secret)


feishu_ready = False
feishu_err = None
try:
    _client = get_feishu_client()
    # 试一次 token 拿不通报错
    _ = _client._get_token()
    feishu_ready = True
except Exception as e:
    feishu_err = str(e)

if feishu_ready:
    st.success("✓ 飞书已连接")
else:
    st.error(f"❌ 飞书连接失败: {feishu_err}")
    st.info("如果是部署在 Streamlit Cloud → 右下角 ⋮ → Settings → Secrets, "
            "添加 `FEISHU_APP_ID` 和 `FEISHU_APP_SECRET` 两行 (按 TOML 格式)。")

# ---------------- 上传 + 参数 ----------------
uploaded = st.file_uploader(
    "选择工厂出货单 (.xlsx)",
    type=['xlsx'],
    help="支持: 雅希(广州) / 倾诚(二厂) / 黛宝(三厂) / 猛哥(四厂)"
)

factory_label = st.selectbox(
    "工厂",
    ["自动识别", "A 雅希 (广州)", "B 倾诚 (二厂)", "D 黛宝 (三厂)", "E 猛哥 (四厂)"],
)

col3, col4 = st.columns(2)
with col3:
    pt = st.number_input("铂金价", value=380.0, step=1.0, format="%.2f")
with col4:
    au = st.number_input("黄金价", value=900.0, step=1.0, format="%.2f")

# 天然钻查飞书 GIA 库存的窗口: 0 = 查全部 sheet (慢一些但稳, 客户几个月前配的老石头也能捞到)
GIA_MONTHS = 0

# v15: 对齐终端 process.py
# - 工厂单 _完成.xlsx 永远生成 (基础产物)
# - 有客户单 + 飞书连通 → 自动查飞书 (要写入完成文件的客户名 / 利润)
# - 现货件永远包含
include_spots = True
use_feishu = feishu_ready   # 隐式: 有客户单就查 (不再让用户选)

st.divider()
st.markdown("**附加产物**（可选, 不勾也会生成工厂单完成文件）:")

col_a, col_b = st.columns(2)
with col_a:
    do_jst = st.checkbox(
        "📦 生成聚水潭入库 Excel",
        value=True,
    )
with col_b:
    sync_feishu = st.checkbox(
        "✏️ 同步今日成本到飞书",
        value=False,
        disabled=not feishu_ready,
        help="勾上后写入飞书「镶嵌成本」字段",
    )

overwrite = False
debug_mode = False
if sync_feishu:
    st.warning(
        "⚠️ **将写入飞书**: 修改「镶嵌成本」字段。"
        "默认叠加 (原值+今天值), 已退款件自动写 0。"
        "**确认这是今天的最新账单**再生成。"
    )
    col_o, col_d = st.columns(2)
    with col_o:
        overwrite = st.checkbox(
            "覆盖模式 (不叠加, 直接覆盖)",
            value=False,
        )
    with col_d:
        debug_mode = st.checkbox(
            "🔬 诊断模式 (写入异常时打开)",
            value=False,
            help="详情行附加 record_id / 写入 payload / 飞书返回 code",
        )


# ---------------- 工厂识别 ----------------
FACTORY_HINTS = [
    ('郑国远', 'A'), ('雅希', 'A'), ('广州', 'A'),
    ('倾城', 'B'), ('倾诚', 'B'), ('JC', 'B'),
    ('真诚出货单', 'B'),   # 二厂 天然钻单 (真诚部门, 文件名唯一识别)
    ('郑总', 'E'), ('天然钻石', 'E'),
    ('008-', 'D'), ('SG2026', 'D'), ('-SG', 'D'), ('黛宝', 'D'),
    ('ZC', 'BUXIN'),        # 布心 (培育 ZC2026年出货 / 天然 ZC26年出货, 靠内容分辨天然/培育)
]


def auto_detect_factory(filename):
    name = os.path.basename(filename)
    for keyword, code in FACTORY_HINTS:
        if keyword in name:
            return code
    return None


def detect_default_material(filename):
    name = os.path.basename(filename).lower()
    if 'pt950' in name or 'pt952' in name or re.search(r'pt\s?95[02]', name):
        return 'PT950'
    return None


def detect_is_natural(filename):
    """文件名识别天然钻工厂单:
       - 猛哥: '天然钻石' / '天然钻'
       - 二厂: '真诚出货单'
       黛宝 / 布心的文件名跟培育版几乎一样, 用 detect_is_natural_by_content 内容识别兜底.
    """
    name = os.path.basename(filename)
    return ('天然钻石' in name or '天然钻' in name
            or '真诚出货单' in name)


def detect_is_natural_by_content(xlsx_path):
    """打开 xlsx 看 sheet 名, 有 '结料' / '真诚' 字样 → 天然钻单.
       布心天然 sheet 名 = 'PT出货单 结料 7-4' (培育没'结料'),
       黛宝天然 sheet 名可能含 '真诚' (待用户确认).
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        for sn in wb.sheetnames:
            if '结料' in sn or '真诚' in sn:
                return True
    except Exception:
        pass
    return False


def detect_natural_factory_name(code, filename):
    """把 A/B/D/E code 映射到 natural.PARSERS 里的工厂名.
       code 可能是 'BUXIN' (布心走文件名映射), 或 'B'/'D'/'E' (走 NATURAL_FACTORY_MAP).
    """
    if code == 'BUXIN':
        return '布心'
    return NATURAL_FACTORY_MAP.get(code)


# ---------------- 飞书 GIA 电子表格 客户端 (只做天然钻查询) ----------------
@st.cache_resource
def get_gia_client():
    app_id, app_secret = load_credentials()
    return natural.FeishuSheetClient(app_id, app_secret)


# ---------------- 天然钻流程 (共用: 黛宝 / 布心 / 猛哥 / 二厂) ----------------
def run_natural_workflow(in_path, uploaded_name, factory_name, pt, au, gia_months=6):
    """
    跑一次天然钻流程: 解析 → 查 GIA → 生成聚水潭 xlsx + 下载按钮.
    factory_name: natural.PARSERS 的键 ('黛宝'/'布心'/'猛哥'/'二厂')
    """
    st.subheader(f"💎 天然钻流程 — {factory_name}")
    parser = natural.PARSERS[factory_name]
    with st.spinner(f"解析 {factory_name} 天然钻单..."):
        try:
            items = parser(in_path, au_price=au, pt_price=pt)
        except Exception as e:
            st.error(f"❌ 解析失败: {e}")
            with st.expander("详细错误"):
                st.code(traceback.format_exc())
            return 0

    st.info(f"识别到 **{len(items)}** 件天然钻客订")
    if not items:
        st.warning("⚠️ 没有天然钻客订件, 天然聚水潭跳过")
        return 0

    with st.expander(f"📋 {len(items)} 件客订清单"):
        for it in items:
            st.text(f"  #{it['no']} 客户={it['款号']} 品名={it['品名']} "
                    f"成色={it['材质颜色']} 镶嵌成本={it['镶嵌成本']}")

    # 查 GIA 库存
    st.markdown("**🔍 查飞书 GIA 库存**")
    try:
        gia_client = get_gia_client()
        all_sheets = gia_client.list_sheets(natural.GIA_TOKEN)
        order_sheets_all = [s for s in all_sheets if '订货' in s.get('title', '')]
        # v19.5: 只查当年 sheet (title 以 "26" 结尾 = 2026 年)
        order_sheets = natural._current_year_gia_sheets(order_sheets_all)
        st.caption(f"总共 {len(order_sheets_all)} 个订货 sheet, "
                   f"只查当年 ({datetime.now().year}) → {len(order_sheets)} 个: "
                   + ' / '.join(s.get('title', '') for s in order_sheets))
    except Exception as e:
        st.error(f"❌ GIA 库存查询初始化失败: {e}")
        return 0

    # v19.4: 预加载所有 sheet 的 layout, 让用户看到每 sheet 表头识别到了哪些字段
    with st.expander("🔬 各 sheet 表头识别 (点开看诊断)", expanded=False):
        layout_lines = []
        for sh in order_sheets:
            title = sh.get('title', '')
            try:
                rows, layout = natural._load_gia_sheet(gia_client, sh.get('sheet_id'))
                keys_str = ', '.join(sorted(layout.keys())) if layout else '(空 — 表头未识别)'
                layout_lines.append(f"[{title}] 行数={len(rows)} 识别字段: {keys_str}")
            except Exception as e:
                layout_lines.append(f"[{title}] ❌ 拉取失败: {e}")
        st.code('\n'.join(layout_lines), language=None)

    placeholder = st.empty()
    progress = st.progress(0.0)
    lines = []
    for i, it in enumerate(items):
        gia = natural.search_gia(gia_client, order_sheets, it['款号'], it['证书号'])
        it['_gia'] = gia
        attrs = gia.get('attrs') or {}
        red_mark = ' 🔴多散货' if gia['散货行数'] >= 2 else ''
        hit_mark = '✓' if attrs.get('证书') else '✗'
        lines.append(
            f"  {hit_mark} #{it['no']} {it['款号']} (证书号={it['证书号'] or '空'}) → "
            f"成本1={gia['cost1']:.0f} 成本2={gia['cost2']:.0f} "
            f"(散货{gia['散货行数']}条){red_mark}"
        )
        if attrs.get('证书'):
            lines.append(
                f"      属性: {attrs.get('证书','')} {attrs.get('形状','')} "
                f"{attrs.get('主石重量','')}ct {attrs.get('颜色等级','')} "
                f"{attrs.get('净度','')}"
            )
        # v19.4: 未匹配上的 → 打印 debug 详情帮定位
        if not attrs.get('证书'):
            lines.append(f"      🔍 {gia.get('debug', '(无 debug)')[:300]}")
        placeholder.code('\n'.join(lines), language=None)
        progress.progress((i + 1) / len(items))
    progress.empty()

    # 生成 xlsx
    jst_rows = [natural.build_jst_row(it, it['_gia'], factory_name) for it in items]
    out_path = tempfile.mktemp(suffix='.xlsx')
    date_str = datetime.now().strftime('%m-%d')
    natural.gen_jst_xlsx(jst_rows, out_path, sheet_title=f'{factory_name}天然{date_str}')
    with open(out_path, 'rb') as f:
        data = f.read()
    try: os.unlink(out_path)
    except OSError: pass

    red_count = sum(1 for r in jst_rows if r.get('_红底'))
    if red_count:
        st.warning(f"⚠️ {red_count} 件多散货, 客户名已标红 — 请人工核验")

    fname = f'聚水潭_{factory_name}天然_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    st.download_button(
        f"📥 下载 {factory_name} 天然钻聚水潭 ({len(items)} 件)",
        data=data, file_name=fname,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type="primary",
        key=f'dl_nat_{factory_name}_{datetime.now().timestamp()}',
    )
    st.session_state.setdefault('history', []).append({
        '时间': datetime.now().strftime('%H:%M:%S'),
        '类型': f'聚水潭-{factory_name}天然',
        '工厂': factory_name,
        '件数': len(items),
        '文件名': fname,
        '_data': data,
    })
    return len(items)


# v13: 天然钻石按文件名自动识别 (猛哥单子文件名都会写)
is_natural = False
if uploaded is not None:
    is_natural = detect_is_natural(uploaded.name)
    if is_natural:
        st.info("💎 文件名含「天然钻石」→ 主石类别会写'天然钻石', 商品名加'天然'前缀")

# ---------------- 主流程 ----------------
# v15: 只要有上传文件就能跑 (即使两个附加任务都没勾, 工厂单完成文件也会生成)
if st.button("🚀 开始", disabled=uploaded is None, type="primary"):
    try:
        suffix = '.xlsx'
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
            f.write(uploaded.read())
            in_path = f.name

        # 识别工厂
        if factory_label.startswith("自动"):
            code = auto_detect_factory(uploaded.name)
            if not code:
                st.error(f"❌ 文件名 `{uploaded.name}` 无法自动识别工厂，请上方手动选择")
                os.unlink(in_path)
                st.stop()
        else:
            code = factory_label[0]

        # v19: 天然钻兜底识别 (文件名兜不住的用内容 - sheet 名含"结料"/"真诚")
        if not is_natural:
            try:
                if detect_is_natural_by_content(in_path):
                    is_natural = True
                    st.info("💎 检测到 sheet 名含「结料」/「真诚」→ 自动切换到天然钻流程")
            except Exception:
                pass
        # BUXIN 是布心天然专用 code (培育钻网站没有布心工厂)
        if code == 'BUXIN':
            is_natural = True

        st.info(f"🏭 工厂: **{code}** | 铂 {pt} | 金 {au} | "
                f"{'天然' if is_natural else '培育'}钻石 | "
                f"{'查飞书' if use_feishu else '不查飞书'}")

        # v19: 单纯天然钻工厂 (布心 / 黛宝 / 二厂) → 只跑天然流程, 跳过培育
        if is_natural and code != 'E':
            natural_factory = detect_natural_factory_name(code, uploaded.name)
            if not natural_factory:
                st.error(f"❌ 无法定位天然钻 parser (code={code}, 文件={uploaded.name})")
                try: os.unlink(in_path)
                except OSError: pass
                st.stop()
            run_natural_workflow(in_path, uploaded.name, natural_factory, pt, au,
                                 gia_months=GIA_MONTHS)
            try: os.unlink(in_path)
            except OSError: pass
            st.stop()

        default_material = detect_default_material(uploaded.name)
        parse_kwargs = dict(pt_price=pt, au_price=au)
        if default_material:
            parse_kwargs['default_material'] = default_material

        with st.spinner("解析工厂账单..."):
            items = factories.get_parser(code)(in_path, **parse_kwargs)

        clients = [it for it in items if it['类别'] == '客户单']
        spots = [it for it in items if it['类别'] == '现货']
        repairs = [it for it in items if it['类别'] == '修理']

        st.success(f"✓ 解析 {len(items)} 件 — 客户单 {len(clients)} | 现货 {len(spots)} | 修理 {len(repairs)}")

        # ============== v15: 飞书查询 + 同步 (有客户单自动查) ==============
        feishu_hit = 0
        feishu_miss = []
        write_ok = refunded = accumulated = 0
        write_fails = []
        detail_lines = []
        need_feishu = clients and feishu_ready   # 有客户单 + 飞书连通就查

        if need_feishu:
            label = '查飞书' + (' + 写镶嵌成本' if sync_feishu else ' (只读, 不写)')
            st.subheader(f"🔍 Step: {label}")
            st.caption("✓ 正常 / ⚠ 异常需关注 / ✗ 失败")
            placeholder = st.empty()
            progress = st.progress(0.0, "处理中...")

            def append(line):
                detail_lines.append(line)
                placeholder.code('\n'.join(detail_lines), language=None)

            for i, c in enumerate(clients):
                key = c.get('飞书匹配键')
                if not key:
                    feishu_miss.append((c['no'], '无匹配键'))
                    append(f"  ✗ #{c['no']} 无飞书匹配键")
                    progress.progress((i + 1) / len(clients))
                    continue
                # 查
                try:
                    rec = _client.find_by_order_number(APP_TOKEN, TABLE_ID, key)
                    if not rec and c.get('证书编号'):
                        rec = _client.find_by_cert(APP_TOKEN, TABLE_ID, str(c['证书编号']).strip())
                except Exception as e:
                    feishu_miss.append((c['no'], f'API 错: {e}'))
                    append(f"  ✗ #{c['no']} {key}: API 错 {e}")
                    progress.progress((i + 1) / len(clients))
                    continue
                if not rec:
                    feishu_miss.append((c['no'], f'未找到 ({key})'))
                    append(f"  ✗ #{c['no']} {key} → 飞书找不到")
                    progress.progress((i + 1) / len(clients))
                    continue

                c['_record_id'] = rec['record_id']
                fields = rec['fields']
                customer = _client.get_text(fields.get('客户名称'))
                c['飞书客户名'] = customer
                c['飞书证书编码'] = _client.get_text(fields.get('证书编码'))
                # v14.5: 保留 None 区别 (不再 or 0), 否则裸钻=0 初始化逻辑永远不触发
                c['飞书裸钻成本'] = _client.get_number(fields.get('裸钻成本'))
                c['飞书配石成本'] = _client.get_number(fields.get('配石成本'))
                c['飞书主石'] = _client.get_text(fields.get('主石'))
                c['飞书圈号'] = _client.get_text(fields.get('圈号'))
                c['飞书利润'] = _client.get_number(fields.get('利润'))
                c['飞书利润率'] = _client.get_number(fields.get('利润率'))
                existing_cost = _client.get_number(fields.get('镶嵌成本')) or 0
                c['_飞书原成本'] = existing_cost
                status_val = (fields.get('货物状态')
                              or fields.get('货品状态')
                              or fields.get('状态'))
                status_text = _client.get_text(status_val) or ''
                c['_飞书货物状态'] = status_text
                is_refunded = '已退款' in status_text
                feishu_hit += 1

                # 写
                today_cost = c['镶嵌成本']
                final_cost = existing_cost
                note = ''
                write_failed = False
                res = None
                update_fields = {}

                if sync_feishu:
                    if is_refunded:
                        c['飞书客户名'] = f"{customer or '客户'}已退款做现货"
                        final_cost = 0
                        update_fields = {'镶嵌成本': 0}
                        refunded += 1
                        note = ' [已退款→¥0]'
                    else:
                        if not overwrite and existing_cost > 0:
                            final_cost = round(existing_cost + today_cost)
                            accumulated += 1
                            note = f' (原{int(existing_cost)}+{today_cost})'
                        else:
                            final_cost = today_cost
                        update_fields = {'镶嵌成本': final_cost}
                        if c.get('飞书裸钻成本') is None:
                            update_fields['裸钻成本'] = 0

                    try:
                        res = _client.update_record(APP_TOKEN, TABLE_ID,
                                                    c['_record_id'], update_fields)
                        if res.get('code') == 0:
                            write_ok += 1
                            if not is_refunded:
                                # v14.2: 轮询飞书直到公式刷新拿最新利润
                                fresh, fresh_ok = _fetch_after_update(
                                    _client, key, final_cost, max_wait=4)
                                if fresh:
                                    c['飞书利润'] = _client.get_number(
                                        fresh['fields'].get('利润'))
                                    c['飞书利润率'] = _client.get_number(
                                        fresh['fields'].get('利润率'))
                                if not fresh_ok:
                                    note += ' [⚠️公式刷新慢]'
                        else:
                            write_fails.append((c['no'], str(res)[:60]))
                            write_failed = True
                            note += f' [写失败 {str(res)[:30]}]'
                    except Exception as e:
                        write_fails.append((c['no'], str(e)[:60]))
                        write_failed = True
                        note += f' [写异常 {str(e)[:30]}]'

                # 详情行
                profit = c.get('飞书利润')
                rate = c.get('飞书利润率')
                profit_str = str(int(profit)) if isinstance(profit, (int, float)) else '?'
                rate_str = f"{rate*100:.1f}%" if isinstance(rate, (int, float)) else '?'

                # 异常标识: 写失败 / 已退款 / 利润率 < 15% 或 > 70% (这两阈值你可改)
                mark = '✓'
                if write_failed:
                    mark = '✗'
                elif is_refunded:
                    mark = '⚠'
                elif isinstance(rate, (int, float)) and (rate < 0.15 or rate > 0.70):
                    mark = '⚠'

                cost_part = f"¥{final_cost}" if sync_feishu else f"飞书¥{int(existing_cost)}"
                line = (f"  {mark} #{c['no']} {key} → {customer or '?'}  "
                        f"{cost_part}{note}  利润={profit_str} 利润率={rate_str}")
                if debug_mode and sync_feishu:
                    # 诊断: 显示 record_id + 写入 payload + 飞书返回
                    rec_id = c.get('_record_id', '?')[-8:] if c.get('_record_id') else '?'
                    payload = update_fields if sync_feishu else {}
                    res_short = res.get('code') if isinstance(res, dict) else '?'
                    res_msg = res.get('msg', '')[:30] if isinstance(res, dict) else ''
                    line += f"\n      🔬 rid=…{rec_id} payload={payload} → code={res_short} msg={res_msg}"
                append(line)
                progress.progress((i + 1) / len(clients))

            progress.empty()
            # 汇总
            summary = f"✓ 飞书匹配 {feishu_hit}/{len(clients)}"
            if sync_feishu:
                summary += f" | 写入 {write_ok}"
                if accumulated: summary += f" | 叠加 {accumulated}"
                if refunded: summary += f" | 已退款 {refunded}"
                if write_fails: summary += f" | 失败 {len(write_fails)}"
            st.success(summary)

        # ============== v13: 生成入库 Excel (只在勾选时) ==============
        if do_jst:
            st.subheader("📦 生成聚水潭入库 Excel")
            targets = list(clients) + list(spots)   # v13: 永远含现货件
            if not targets:
                st.warning("⚠️ 没有可入库的件 (客户单/现货都为 0)")
            else:
                rows = []
                for it in targets:
                    row = jst.build_row_from_item(
                        item=it,
                        factory_code=code,
                        feishu_cert=it.get('飞书证书编码') or it.get('证书编号'),
                        feishu_luozuan=it.get('飞书裸钻成本') or 0,
                        feishu_peishi=it.get('飞书配石成本') or 0,
                        feishu_main_stone=it.get('飞书主石'),
                        feishu_ring_size=it.get('飞书圈号'),
                        total_weight=it.get('总重'),
                        is_natural=is_natural,
                    )
                    rows.append(row)

                out_path = tempfile.mktemp(suffix='.xlsx')
                added, _ = jst.generate_or_append(out_path, rows)

                st.success(f"✅ 已生成 {added} 行")

                with st.expander("📋 预览前 8 行"):
                    preview = []
                    for r in rows[:8]:
                        preview.append({
                            '商品编码': r.get('商品编码'),
                            '商品名': r.get('商品名称'),
                            '成色': r.get('成色'),
                            '主石类别': r.get('主石类别'),
                            '主石ct': r.get('主石重量'),
                            '颜色': r.get('颜色等级'),
                            '净度': r.get('净度'),
                            '圈号': r.get('指圈号'),
                            '总重': r.get('总重'),
                            '成本': r.get('成本价'),
                        })
                    st.dataframe(preview, use_container_width=True, hide_index=True)

                with open(out_path, 'rb') as f:
                    jst_data = f.read()
                jst_fname = f'聚水潭入库_{code}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                st.download_button(
                    "📥 下载聚水潭入库 Excel", data=jst_data, file_name=jst_fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type="primary",
                    key=f'dl_jst_{datetime.now().timestamp()}',
                )
                # session 历史 (聚水潭)
                st.session_state.setdefault('history', []).append({
                    '时间': datetime.now().strftime('%H:%M:%S'),
                    '类型': '聚水潭入库',
                    '工厂': code,
                    '件数': added,
                    '文件名': jst_fname,
                    '_data': jst_data,
                })

                try:
                    os.unlink(out_path)
                except OSError:
                    pass

        # ============== v15: 工厂单 _完成.xlsx 永远生成 (对齐终端) ==============
        st.subheader("📑 工厂单完成文件")
        try:
            writer = factories.get_writer(code)
            done_path = tempfile.mktemp(suffix='.xlsx')
            writer(in_path, items, done_path)
            with open(done_path, 'rb') as f:
                done_data = f.read()
            # v15.3: 文件名格式 = 日期 + 原始文件名 (例 20260630郑总-6月份18k-培育钻(44).xlsx)
            base = uploaded.name.rsplit('.', 1)[0]
            done_fname = f'{datetime.now().strftime("%Y%m%d")}{base}.xlsx'
            st.download_button(
                "📥 下载工厂单 _完成.xlsx",
                data=done_data, file_name=done_fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f'dl_done_{datetime.now().timestamp()}',
                type="primary",
            )
            cap_parts = []
            if clients: cap_parts.append("客户名/利润/利润率")
            if spots: cap_parts.append("现货 AD 单价成本")
            st.caption(f"含 {' + '.join(cap_parts)}" if cap_parts else "已格式化")
            st.session_state.setdefault('history', []).append({
                '时间': datetime.now().strftime('%H:%M:%S'),
                '类型': '工厂单完成',
                '工厂': code,
                '件数': len(items),
                '文件名': done_fname,
                '_data': done_data,
            })
            try:
                os.unlink(done_path)
            except OSError:
                pass
        except Exception as e:
            st.error(f"生成工厂单完成文件失败: {e}")
            with st.expander("详细错误"):
                st.code(traceback.format_exc())

        # v19: 猛哥天然钻单同表混合 → 培育钻流程跑完后, 追加天然钻流程 (19楼)
        if is_natural and code == 'E':
            st.divider()
            st.subheader("💎💎 猛哥双份 — 追加天然钻流程 (19楼真诚部门)")
            try:
                run_natural_workflow(in_path, uploaded.name, '猛哥', pt, au,
                                     gia_months=GIA_MONTHS)
            except Exception as e:
                st.error(f"❌ 天然钻流程失败: {e}")
                with st.expander("详细错误"):
                    st.code(traceback.format_exc())

        try:
            os.unlink(in_path)
        except OSError:
            pass

    except Exception as e:
        st.error(f"❌ 出错: {e}")
        with st.expander("详细错误"):
            st.code(traceback.format_exc())

# ---------------- 本次会话历史 ----------------
st.divider()
hist = st.session_state.get('history', [])
if hist:
    st.subheader(f"📚 本次会话生成记录 ({len(hist)} 份)")
    st.caption("⚠️ 浏览器关掉就清空, 重要的请存到本地")
    for idx, h in enumerate(reversed(hist)):
        col_a, col_b, col_c, col_d = st.columns([1.5, 2, 1, 2.5])
        with col_a:
            st.text(f"🕐 {h['时间']}")
        with col_b:
            st.text(f"{h['类型']} ({h['工厂']})")
        with col_c:
            st.text(f"{h['件数']} 件")
        with col_d:
            st.download_button(
                "📥 重下载",
                data=h['_data'],
                file_name=h['文件名'],
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f'redown_{idx}_{h["时间"]}',
            )

# ---------------- 说明 ----------------
st.divider()
with st.expander("ℹ️ 使用说明"):
    st.markdown("""
**步骤**: 上传工厂账单 (.xlsx) → 输入金价 → 「生成」→ 下载入库 Excel

**自动识别工厂关键词**:
- A 雅希: 含「郑国远」「雅希」「广州」
- B 倾诚: 含「倾城」「倾诚」「JC」
- D 黛宝: 含「黛宝」「008-」「SG2026」
- E 猛哥: 含「郑总」「天然钻石」

**查飞书** 勾选后, 系统会根据下单单号 / 证书编号去飞书读:
- 客户名称 → 商品名 (没找到时商品名是品类)
- 圈号 → 指圈号 (只戒指带)
- 主石 → 主石重量 / 颜色 / 净度
- 裸钻成本 / 配石成本 → 成本1

**.xls 文件**: Streamlit Cloud 上没装 libreoffice, 请先另存为 .xlsx

**✏️ 同步飞书镶嵌成本** (v12):
- 默认**不勾**, 只读飞书不写
- 勾上后会把今天工厂账单算的成本写到飞书「镶嵌成本」字段
- 默认**叠加** (原值 + 今天值), 用于同一单分多次出货的场景
- 想直接替换原值就勾「覆盖模式」
- 飞书"货物状态"含"已退款"的件 → 自动写 0 + 客户名加"已退款做现货"
    """)
