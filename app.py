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
from feishu_client import FeishuClient, APP_TOKEN, TABLE_ID, load_credentials


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
            base = uploaded.name.rsplit('.', 1)[0]
            done_fname = f'{base}_完成_{datetime.now().strftime("%H%M%S")}.xlsx'
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
