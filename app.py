"""聚水潭入库 Excel 生成 — 网站版 v11
- 接飞书读: 客户名 / 利润 / 圈号 / 主石 / 裸钻成本 / 配石成本 / 证书编码
- 不写飞书 (网站默认只读, 避免误更)
- 凭证: Streamlit Cloud Secrets 加密存
"""
import streamlit as st
import tempfile, os, traceback, re
from datetime import datetime

import factories
import jushuitan_import as jst
from feishu_client import FeishuClient, APP_TOKEN, TABLE_ID, load_credentials


st.set_page_config(page_title="聚水潭入库生成", page_icon="💎", layout="centered")

st.title("💎 聚水潭入库 Excel 生成")
st.caption("上传工厂出货单 → 查飞书 → 生成聚水潭批量入库模板")

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
    pt = st.number_input("铂金价 PT950 (元/g)", value=380.0, step=1.0, format="%.2f")
with col4:
    au = st.number_input("黄金价 18K (元/g)", value=900.0, step=1.0, format="%.2f")

# 永远包含现货件 (v13 移除选项)
include_spots = True
# 查飞书: 飞书连通就默认勾上
use_feishu = st.checkbox(
    "查飞书补客户名/利润/圈号/主石",
    value=feishu_ready,
    disabled=not feishu_ready,
)

st.divider()
st.markdown("**🎯 任务**（至少选一项）:")

# v13: 两个独立任务
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
if sync_feishu:
    st.warning(
        "⚠️ **将写入飞书**: 修改「镶嵌成本」字段。"
        "默认叠加 (原值+今天值), 已退款件自动写 0。"
        "**确认这是今天的最新账单**再生成。"
    )
    overwrite = st.checkbox(
        "覆盖模式 (飞书已有成本时直接覆盖, 不叠加)",
        value=False,
    )


# ---------------- 工厂识别 ----------------
FACTORY_HINTS = [
    ('郑国远', 'A'), ('雅希', 'A'), ('广州', 'A'),
    ('倾城', 'B'), ('倾诚', 'B'), ('JC', 'B'),
    ('郑总', 'E'), ('天然钻石', 'E'),
    ('008-', 'D'), ('SG2026', 'D'), ('-SG', 'D'), ('黛宝', 'D'),
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
    name = os.path.basename(filename)
    return '天然钻石' in name or '天然钻' in name


# v13: 天然钻石按文件名自动识别 (猛哥单子文件名都会写)
is_natural = False
if uploaded is not None:
    is_natural = detect_is_natural(uploaded.name)
    if is_natural:
        st.info("💎 文件名含「天然钻石」→ 主石类别会写'天然钻石', 商品名加'天然'前缀")

# ---------------- 主流程 ----------------
# 必须勾至少一项任务才能点
can_run = uploaded is not None and (do_jst or sync_feishu)
btn_label = "🚀 开始"
if do_jst and sync_feishu:
    btn_label = "🚀 生成入库 Excel + 同步飞书"
elif do_jst:
    btn_label = "🚀 生成聚水潭入库 Excel"
elif sync_feishu:
    btn_label = "🚀 同步成本到飞书"

if st.button(btn_label, disabled=not can_run, type="primary"):
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

        st.info(f"🏭 工厂: **{code}** | 铂 {pt} | 金 {au} | "
                f"{'天然' if is_natural else '培育'}钻石 | "
                f"{'查飞书' if use_feishu else '不查飞书'}")

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

        # ============== 查飞书 ==============
        feishu_hit = 0
        feishu_miss = []
        if use_feishu and clients and feishu_ready:
            with st.spinner(f"查飞书 ({len(clients)} 个客户单)..."):
                progress = st.progress(0.0, "查询中...")
                for i, c in enumerate(clients):
                    key = c.get('飞书匹配键')
                    if not key:
                        feishu_miss.append((c['no'], '无匹配键'))
                        continue
                    try:
                        rec = _client.find_by_order_number(APP_TOKEN, TABLE_ID, key)
                        if not rec and c.get('证书编号'):
                            rec = _client.find_by_cert(APP_TOKEN, TABLE_ID, str(c['证书编号']).strip())
                    except Exception as e:
                        feishu_miss.append((c['no'], f'API 错: {e}'))
                        continue
                    if not rec:
                        feishu_miss.append((c['no'], f'未找到 ({key})'))
                        continue
                    c['_record_id'] = rec['record_id']
                    fields = rec['fields']
                    c['飞书客户名'] = _client.get_text(fields.get('客户名称'))
                    c['飞书证书编码'] = _client.get_text(fields.get('证书编码'))
                    c['飞书裸钻成本'] = _client.get_number(fields.get('裸钻成本')) or 0
                    c['飞书配石成本'] = _client.get_number(fields.get('配石成本')) or 0
                    c['飞书主石'] = _client.get_text(fields.get('主石'))
                    c['飞书圈号'] = _client.get_text(fields.get('圈号'))
                    c['飞书利润'] = _client.get_number(fields.get('利润'))
                    c['飞书利润率'] = _client.get_number(fields.get('利润率'))
                    c['_飞书原成本'] = _client.get_number(fields.get('镶嵌成本')) or 0
                    # v12: 货物状态 (已退款分支)
                    status_val = (fields.get('货物状态')
                                  or fields.get('货品状态')
                                  or fields.get('状态'))
                    c['_飞书货物状态'] = _client.get_text(status_val) or ''
                    feishu_hit += 1
                    progress.progress((i + 1) / len(clients), f"{i + 1}/{len(clients)}")
                progress.empty()
            st.success(f"✓ 飞书匹配 {feishu_hit}/{len(clients)}")
            if feishu_miss:
                with st.expander(f"⚠️ {len(feishu_miss)} 件未匹配"):
                    for no, why in feishu_miss:
                        st.text(f"  #{no}: {why}")

        # ============== v12: 同步成本到飞书 ==============
        if sync_feishu and clients and feishu_ready:
            to_write = [c for c in clients if c.get('_record_id')]
            if to_write:
                st.subheader("✏️ 同步成本到飞书")
                write_ok = 0
                refunded = 0
                accumulated = 0
                write_fails = []
                with st.spinner(f"写飞书 ({len(to_write)} 件)..."):
                    progress2 = st.progress(0.0, "写入中...")
                    for i, c in enumerate(to_write):
                        today_cost = c['镶嵌成本']
                        status_text = c.get('_飞书货物状态', '')
                        is_refunded = '已退款' in status_text
                        existing_cost = c.get('_飞书原成本') or 0

                        if is_refunded:
                            # 已退款 → 写 0, 客户名改"X已退款做现货"
                            customer = c.get('飞书客户名')
                            c['飞书客户名'] = f"{customer or '客户'}已退款做现货"
                            new_cost = 0
                            update_fields = {'镶嵌成本': 0}
                            refunded += 1
                            note = '[已退款→¥0]'
                        else:
                            if not overwrite and existing_cost > 0:
                                new_cost = round(existing_cost + today_cost)
                                accumulated += 1
                                note = f'(原{int(existing_cost)}+{today_cost})'
                            else:
                                new_cost = today_cost
                                note = ''
                            update_fields = {'镶嵌成本': new_cost}
                            # 裸钻成本空时初始化 0 (跟原 process.py 一致)
                            if c.get('飞书裸钻成本') is None:
                                update_fields['裸钻成本'] = 0

                        try:
                            res = _client.update_record(
                                APP_TOKEN, TABLE_ID, c['_record_id'], update_fields)
                            if res.get('code') == 0:
                                write_ok += 1
                                # 取最新公式刷新值
                                if not is_refunded:
                                    fresh_rec = _client.find_by_order_number(
                                        APP_TOKEN, TABLE_ID, c['飞书匹配键'])
                                    if fresh_rec:
                                        c['飞书利润'] = _client.get_number(
                                            fresh_rec['fields'].get('利润'))
                                        c['飞书利润率'] = _client.get_number(
                                            fresh_rec['fields'].get('利润率'))
                            else:
                                write_fails.append((c['no'], str(res)[:80]))
                        except Exception as e:
                            write_fails.append((c['no'], str(e)[:80]))
                        progress2.progress((i + 1) / len(to_write),
                                           f"{i + 1}/{len(to_write)}")
                    progress2.empty()
                msg = f"✓ 写飞书 {write_ok}/{len(to_write)}"
                if accumulated: msg += f" | 叠加 {accumulated}"
                if refunded: msg += f" | 已退款 {refunded}"
                if write_fails: msg += f" | 失败 {len(write_fails)}"
                st.success(msg)
                if write_fails:
                    with st.expander(f"❌ {len(write_fails)} 件写入失败"):
                        for no, why in write_fails:
                            st.text(f"  #{no}: {why}")

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
                    data = f.read()
                fname = f'聚水潭入库_{code}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                st.download_button(
                    "📥 下载聚水潭入库 Excel", data=data, file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type="primary",
                )

                try:
                    os.unlink(out_path)
                except OSError:
                    pass

        try:
            os.unlink(in_path)
        except OSError:
            pass

    except Exception as e:
        st.error(f"❌ 出错: {e}")
        with st.expander("详细错误"):
            st.code(traceback.format_exc())

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
