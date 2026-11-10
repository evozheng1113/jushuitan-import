"""聚水潭入库 Excel 生成 — 网站版 v11
- 接飞书读: 客户名 / 利润 / 圈号 / 主石 / 裸钻成本 / 配石成本 / 证书编码
- 不写飞书 (网站默认只读, 避免误更)
- 凭证: Streamlit Cloud Secrets 加密存
"""
import streamlit as st
import tempfile, os, traceback, re, time, math, hashlib, uuid
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
st.caption("上传出货单 → 核对金价 → 生成文件或同步成本")
st.caption("版本：2026.11.10 · 流水改为原样导入")

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
    # v24.4: 真诚多维表结构诊断按钮
    with st.expander("连接诊断（遇到问题再打开）"):
        if st.button("拉取真诚多维表字段 + 样本", key='probe_zc'):
            try:
                from zhencheng_bitable import probe_zhencheng_table
                info = probe_zhencheng_table(_client, sample_size=3)
                st.markdown(f"**字段列表** ({len(info['fields'])} 个):")
                for f in info['fields']:
                    st.text(f"  {f['name']:20s}  type={f['type']}  ui_type={f['ui_type']}")
                st.markdown(f"**样本记录** ({len(info['samples'])} 条):")
                if info.get('list_code') and info['list_code'] != 0:
                    st.error(f"list records 失败: code={info['list_code']} msg={info.get('list_msg')}")
                for i, s in enumerate(info['samples'], 1):
                    st.markdown(f"— 样本 {i} record_id=`{s.get('record_id')}`")
                    for k, v in (s.get('fields') or {}).items():
                        st.text(f"    {k}: {repr(v)[:200]}")
            except Exception as e:
                st.error(f"诊断失败: {e}")
                with st.expander("详细错误"):
                    st.code(traceback.format_exc())
else:
    st.error(f"❌ 飞书连接失败: {feishu_err}")
    st.info("如果是部署在 Streamlit Cloud → 右下角 ⋮ → Settings → Secrets, "
            "添加 `FEISHU_APP_ID` 和 `FEISHU_APP_SECRET` 两行 (按 TOML 格式)。")

# ---------------- 上传 + 参数 ----------------
def download_button(*args, **kwargs):
    """点下载不重跑页面，否则会打断正在写飞书的任务。"""
    return st.download_button(*args, on_click="ignore", **kwargs)


class ConvertedUpload:
    """.xls 转成 .xlsx 后，让后面的流程照常使用。"""
    def __init__(self, name, data):
        self.name, self.size, self._data = name, len(data), data
    def getvalue(self):
        return self._data
    def read(self):
        return self._data


def load_upload(file):
    if file is None or not file.name.lower().endswith('.xls'):
        return file
    from xls_convert import xls_to_xlsx
    cache = st.session_state.setdefault('_xls_cache', {})
    raw = file.getvalue()
    key = hashlib.sha256(raw).hexdigest()
    if key not in cache:
        cache[key] = xls_to_xlsx(raw)
    return ConvertedUpload(file.name.rsplit('.', 1)[0] + '.xlsx', cache[key])


uploaded = st.file_uploader(
    "选择工厂出货单 (.xlsx / .xls)",
    type=['xlsx', 'xls'],
    help="支持: 雅希(广州) / 倾诚(二厂) / 黛宝(三厂) / 猛哥(四厂) / 布心(天然钻)"
)
try:
    uploaded = load_upload(uploaded)
except Exception as exc:
    st.error(f"无法读取这份 .xls：{exc}。请用 Excel/WPS 另存为 .xlsx 后再上传。")
    st.stop()

archive_repeat_key = None
# 文件历史与成本成功状态分开；记录失败不当成首次上传。
if uploaded and feishu_ready:
    upload_key = hashlib.sha256(uploaded.getvalue()).hexdigest()
    upload_cache = st.session_state.setdefault('_upload_checks_v3', {})
    if upload_key not in upload_cache:
        try:
            from upload_registry import check_and_record
            from finished_products import FinishedProductsClient
            upload_id, upload_secret = load_credentials()
            upload_client = FinishedProductsClient(upload_id, upload_secret)
            upload_cache[upload_key] = check_and_record(upload_client, uploaded.getvalue(), uploaded.name)
        except Exception as exc:
            upload_cache[upload_key] = {'unknown': True, 'reason': str(exc)[:500]}
    upload_info = upload_cache[upload_key]
    if upload_info.get('unknown'):
        st.warning("上传记录查询失败，尚无法判断是否上传过。")
        with st.expander("查看失败原因"):
            st.text(upload_info.get('reason') or '请点击重查，获取具体原因。')
        if st.button("重查上传记录"):
            upload_cache.pop(upload_key, None)
            st.rerun()
    elif upload_info.get('record_error'):
        st.warning("已核对，未找到历史上传记录；本次记录保存失败。")
        with st.expander("查看失败原因"):
            st.text(upload_info['record_error'])
        if st.button("重试保存上传记录"):
            upload_cache.pop(upload_key, None)
            st.rerun()
    else:
        st.caption("已加载过" if upload_info.get('seen') else "首次加载")
    action_status = st.empty()
    if not upload_info.get('unknown'):
        from upload_registry import status_lines
        if upload_info.get('action_record_error'):
            action_status.warning("上次操作记录保存失败，请重查处理记录。")
        else:
            action_status.info("\n\n".join(status_lines(upload_info)))
        if upload_info.get('archives'):
            archive_repeat_key = f"archive_repeat_{upload_key}_{upload_info.get('archive_count', 0)}"
            st.checkbox("这批工厂单已写过流水，本次确实要再写一份", value=False, key=archive_repeat_key,
                        help="只影响这一次。请先确认流水里没有这批，或确实需要重复写入。")
    if not upload_info.get('unknown') and not upload_info.get('record_error'):
        if st.button("重查处理记录"):
            upload_cache.pop(upload_key, None)
            st.rerun()

factory_label = st.selectbox(
    "工厂",
    ["自动识别", "A 雅希 (广州)", "B 倾诚 (二厂)", "D 黛宝 (三厂)",
     "E 猛哥 (四厂)", "布心 (天然钻)"],
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
st.markdown("**处理选项**")

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

spots_only = st.checkbox("只补写现货到成品新单（上次被打断时用）", value=False, disabled=not feishu_ready,
                         help="金价要和上次一样。只写现货成本，不动客户单，不写流水。")
if spots_only:
    sync_feishu = False
    st.warning("只会把这份单的现货成本叠加到成品新单。确认上次现货确实没写进去再点开始。")

overwrite = False
debug_mode = False
if sync_feishu:
    overwrite = st.checkbox("覆盖模式 (不叠加, 直接覆盖)", value=False)
    with st.expander("诊断设置"):
        debug_mode = st.checkbox("显示诊断信息", value=False)

buxin_initial = {'PT950': False, '18K': False}
if uploaded and sync_feishu and ('ZC' in uploaded.name.upper() or '布心' in factory_label):
    with st.expander("布心首次接入：核对已有成本"):
        st.caption("仅在首次接入遇到已有成本时使用。分别确认本文件最后一批是否尚未同步；不确定就不勾选。以后自动跳过已同步出货。")
        buxin_initial['PT950'] = st.checkbox("确认 PT 最后一批尚未同步，允许更新已有成本", key='buxin_pt_' + hashlib.sha256(uploaded.getvalue()).hexdigest())
        buxin_initial['18K'] = st.checkbox("确认 18K 最后一批尚未同步，允许更新已有成本", key='buxin_18k_' + hashlib.sha256(uploaded.getvalue()).hexdigest())


def remember_action(event):
    if not uploaded or not feishu_ready:
        return
    info = st.session_state.get('_upload_checks_v3', {}).get(upload_key, {})
    try:
        from upload_registry import record_action, status_lines
        from finished_products import FinishedProductsClient
        log_id, log_secret = load_credentials()
        record_action(FinishedProductsClient(log_id, log_secret), info, uploaded.name, event)
        info.pop('action_record_error', None)
        action_status.info("\n\n".join(status_lines(info)))
    except Exception:
        info['action_record_error'] = True
        action_status.warning("本次操作记录未保存，状态待核对。")


def remember_generated(kind='complete', scope=''):
    # 只有真正生成聚水潭入库文件，才能设置该项状态。
    if kind == 'jst':
        remember_action({'kind':'jst_generated', 'run_id':current_run_id, 'scope':scope})


def track_cost(total, written=0, already=0):
    if cost_progress is not None:
        cost_progress.add(total, written, already)
        remember_action(cost_progress.event())


def track_sheet_cost(result, total):
    from upload_registry import sheet_confirmed
    confirmed = sheet_confirmed(result)
    track_cost(total, len(result.get('matched', [])) if confirmed else 0)
    if result.get('matched') and not confirmed:
        st.warning("成品成本写入未获成功确认，请核对，勿当作已同步。")


def sync_finished_spots(items):
    try:
        from finished_products import (FinishedProductsClient, sync_costs,
                                        build_sync_items_from_factory_items)
        sync_items = build_sync_items_from_factory_items(items)
        if sync_items:
            st.divider()
            st.subheader(f"🔗 同步镶嵌成本到成品新单 ({len(sync_items)} 条)")
            fp_client = FinishedProductsClient(
                *load_credentials())
            with st.spinner(f"正在同步成品成本…"):
                result = sync_costs(fp_client, sync_items, overwrite=overwrite)
            track_sheet_cost(result, len(sync_items))
            if result['matched']:
                remember_action({'kind': 'spot_sync', 'run_id': current_run_id, 'count': len(result['matched'])})
            matched = result['matched']
            unmatched = result['unmatched']
            errors = result['errors']

            from upload_registry import sheet_confirmed
            if matched and sheet_confirmed(result):
                h_updated = sum(1 for m in matched if m.get('new_h') is not None)
                st.success(
                    f"✅ 命中 {len(matched)} 条, M 列已更新"
                    + (f", H 列 {h_updated} 条去掉「待出货」" if h_updated else "")
                )
                with st.expander(f"命中详情 ({len(matched)})", expanded=False):
                    lines = []
                    for m in matched:
                        h_str = ''
                        if m.get('new_h') is not None:
                            h_str = f"  H: {m['old_h']} → {m['new_h']}"
                        p_display = m.get('p_val') or m.get('match_key')
                        lines.append(
                            f"  r{m['row']:>4}  P={p_display:<18} "
                            f"M {m['old_m']:>6} → "
                            f"{m['new_m']:>6}  {m['name']}{h_str}"
                        )
                    st.code('\n'.join(lines), language=None)
            if unmatched:
                st.warning(f"⚠️ {len(unmatched)} 条没匹配到成品新单: "
                           + ', '.join(unmatched[:15])
                           + ('...' if len(unmatched) > 15 else ''))
            if errors:
                for e in errors:
                    st.error(f"❌ {e}")
    except Exception as e:
        track_cost(max(1, len(locals().get("sync_items", []))))
        st.error(f"❌ 同步成品新单失败: {e}")
        with st.expander("详细错误"):
            st.code(traceback.format_exc())


def track_bitable_cost(result, total):
    from upload_registry import already_synced
    track_cost(total, result.get('updated', 0), already_synced(result))



archive_queue = []


def queue_archive(data, source_path, items, factory, filename, flow):
    """流水最慢，排到所有成本写完之后。猛哥双份流程是同一张表，只写一次。"""
    from factory_archive import NAMES
    name = NAMES.get(factory, factory)
    if any(NAMES.get(job[3], job[3]) == name for job in archive_queue):
        return
    archive_queue.append((data, source_path, items, factory, filename, flow))
    st.info("完成单现在就可以下载。工厂流水会在页面最下方写入，写完前请勿关闭或刷新页面。")


def finish_run():
    """成本结果先落定，再写流水；写流水慢或被中断都不影响成本。"""
    if cost_progress is not None:
        remember_action(cost_progress.event(complete=True))
    if archive_queue:
        st.divider()
        st.subheader("📒 写入工厂流水")
    while archive_queue:
        archive_completed(*archive_queue.pop(0))


def archive_completed(data, source_path, items, factory, filename, flow):
    from factory_archive import ArchiveUncertain, LayoutIncomplete, NAMES
    from upload_registry import ARCHIVE_BLOCKING, ARCHIVE_LABELS, archive_state, record_action, status_lines
    name = NAMES.get(factory, factory)
    info = st.session_state.get('_upload_checks_v3', {}).get(upload_key, {}) if uploaded else {}
    registry = None

    rows = None
    bar = None

    def note(state):
        record_action(registry, info, uploaded.name, {'kind': 'archive', 'run_id': current_run_id,
                                                      'factory': name, 'state': state, 'batch': rows.batch})
        action_status.info("\n\n".join(status_lines(info)))

    def finish(state):
        try:
            note(state)
        except Exception:
            # 保留 started 记录，下次按“结果未确认”处理，不会被当成没写过。
            st.warning("流水结果记录保存失败，再次处理这个文件前请先查看工厂流水。")

    try:
        from factory_archive import snapshot, batch_workbook
        from flow_import import import_batch
        from finished_products import FinishedProductsClient
        rows = snapshot(data, source_path, items, factory, filename, flow)
        book = batch_workbook(data, source_path, rows, factory)  # 先在本机做好，出错不会留下半条记录
        app_id, app_secret = load_credentials()
        client = FinishedProductsClient(app_id, app_secret)
        previous = archive_state(client, name, rows.batch)
        repeat = bool(archive_repeat_key and st.session_state.get(archive_repeat_key))
        if previous in ARCHIVE_BLOCKING and not repeat:
            info.setdefault('archives', {})[name] = previous  # 别的页面写过时，让确认框显示出来
            st.warning(f"这批{name}工厂单的流水{ARCHIVE_LABELS[previous]}，本次没有再写入。完成文件仍可下载。"
                       "如确实要再写一份，先核对流水，勾选上方的确认后重新点开始。")
            return
        registry = client
        note('started')  # 先留记录再写：记录失败就不写，中途断开也能被下次发现。
        bar = st.progress(0.0, "正在写入工厂流水，请勿关闭或刷新页面…")

        def progress(done, total):
            bar.progress(done / total, f"正在写入工厂流水 {done}/{total}，请勿关闭或刷新页面…")

        url, notes = import_batch(client, book, name, uploaded.name, len(items), rows.batch, progress)
        finish('written')
        bar.empty()
        st.success(f"已导入工厂流水 · {name}　[打开这批流水]({url})")
        for message in notes:
            st.warning(message)
    except Exception as exc:
        if bar is not None:
            bar.empty()
        if registry is not None and info.get('archives', {}).get(name) == 'started':
            finish('layout_incomplete' if isinstance(exc, LayoutIncomplete)
                   else 'uncertain' if isinstance(exc, ArchiveUncertain) else 'failed')
        if isinstance(exc, LayoutIncomplete):
            st.session_state["archive_pending_layout"] = (exc.rows, factory, exc.first_row)
        st.warning("工厂流水未确认归档，完成文件仍可下载。")
        with st.expander("查看归档原因"):
            st.write(str(exc))

# Repair operates independently of Start and never invokes cost synchronization.
with st.expander("仅修复归档格式"):
    st.caption("用于流水已追加、格式失败的批次。上传当次工厂单完成文件，核对一致后修复原位置。")
    if st.session_state.get("archive_pending_layout") and st.button("修复本页刚才失败的批次"):
        try:
            from factory_archive import TOKEN, repair_layout
            from finished_products import FinishedProductsClient
            pending_rows, pending_factory, pending_start = st.session_state["archive_pending_layout"]
            aid, secret = load_credentials()
            repair_layout(FinishedProductsClient(aid, secret, spreadsheet_token=TOKEN), pending_rows, pending_factory, pending_start)
            del st.session_state["archive_pending_layout"]
            st.success("原位置格式已修复，未追加流水、未写入成本。")
        except Exception as exc:
            st.error(str(exc))
    repair_source = st.file_uploader("原出货单（完成文件含公式时用于读取原计算结果）", type=["xlsx", "xls"], key="archive_repair_source")
    repair_file = st.file_uploader("工厂单完成文件", type=["xlsx"], key="archive_repair_file")
    repair_factory = st.selectbox("归档工厂", ["A", "B", "D", "E", "BUXIN"],
        format_func=lambda v: {"A":"广州", "B":"二厂", "D":"黛宝", "E":"猛哥", "BUXIN":"布心"}[v], key="archive_repair_factory")
    repair_start = st.number_input("提示中的起始行", min_value=1, value=1, step=1, key="archive_repair_start")
    repair_sheets = []
    if repair_file is not None:
        try:
            import openpyxl as repair_xl
            from factory_archive import default_repair_sheets
            from io import BytesIO as RepairBytesIO
            repair_book = repair_xl.load_workbook(RepairBytesIO(repair_file.getvalue()), read_only=True)
            repair_sheets = st.multiselect("本次归档的sheet（保持原顺序）", repair_book.sheetnames, default=default_repair_sheets(repair_book.sheetnames, repair_factory))
            repair_book.close()
        except Exception as exc:
            st.error(f"无法读取完成文件：{exc}")
    if st.button("核对并仅修复格式", disabled=not repair_sheets):
        try:
            from factory_archive import TOKEN, snapshot, repair_layout
            from finished_products import FinishedProductsClient
            raw = repair_file.getvalue()
            repair_source = load_upload(repair_source)
            # Completed files with missing formula caches must be recalculated first;
            # never guess values or mutate a non-matching region.
            repair_rows = snapshot(raw, RepairBytesIO(repair_source.getvalue() if repair_source else raw), [{"_sheet": n} for n in repair_sheets],
                repair_factory, repair_file.name, "格式修复")
            aid, secret = load_credentials()
            target = FinishedProductsClient(aid, secret, spreadsheet_token=TOKEN)
            title = repair_layout(target, repair_rows, repair_factory, int(repair_start))
            st.success(f"已修复{title}原位置格式，未追加流水、未写入成本。")
            if getattr(repair_rows, "format_fallback", False):
                st.caption("部分格式飞书不支持，已保留全部数值和其他格式。")
        except Exception as exc:
            st.error(str(exc))

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
    """打开 xlsx 看 sheet 名判断是否天然钻单:
       - 布心天然: sheet 名含 '结料'   (培育钻布心没有 '结料' 字样)
       - 黛宝天然: sheet 名以 '出货单-' 开头 (培育钻黛宝是 '008-SG*')
       - 通用   : sheet 名含 '真诚'
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        for sn in wb.sheetnames:
            if '结料' in sn or '真诚' in sn:
                return True
            if sn.startswith('出货单-'):
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


def sync_zhencheng_finished_items(items):
    spot_items = [it for it in items if it.get('类别') == '部门-真诚-现货']
    if spot_items:
        try:
            from finished_products import FinishedProductsClient, sync_costs
            app_id, app_secret = load_credentials()
            st.divider()
            st.subheader(f"🔗 真诚现货 (E-ZC-*) → 真诚成品新单 ({len(spot_items)} 件)")

            # 只匹配指定工作表，不回退到培育“成品新单”。
            probe_client = FinishedProductsClient(app_id, app_secret)
            zc_sheet_id = probe_client.find_sheet_id_by_name('真诚成品新单')
            st.caption(f"命中 sheet: 真诚成品新单 (sheet_id=`{zc_sheet_id}`)")

            if zc_sheet_id:
                # 用探测到的 sheet_id 创建 client
                zc_client = FinishedProductsClient(app_id, app_secret, sheet_id=zc_sheet_id)
                # 构造 sync_items: match_key = 款号 (E-ZC-*), cost = 镶嵌成本, name = 品名
                sync_items = [
                    {'match_key': str(it.get('款号') or it.get('单号')).strip(),
                     'cost': it.get('镶嵌成本'),
                     'name': f"#{it.get('no')} {it.get('品名', '')}"}
                    for it in spot_items
                    if (it.get('款号') or it.get('单号')) and it.get('镶嵌成本')
                ]
                with st.spinner("正在同步成品成本…"):
                    zc_res = sync_costs(zc_client, sync_items, cost_col_letter='M', overwrite=overwrite)
                track_sheet_cost(zc_res, len(spot_items))
                st.success(f"✅ 匹配 **{len(zc_res['matched'])}** 条 / 提交 {len(sync_items)} 条")
                if zc_res['unmatched']:
                    st.warning(f"⚠️ {len(zc_res['unmatched'])} 条没匹配到: "
                               f"{', '.join(zc_res['unmatched'][:10])}"
                               + ("..." if len(zc_res['unmatched']) > 10 else ""))
                if zc_res['errors']:
                    st.error(f"❌ {len(zc_res['errors'])} 条错误")
                    for e_ in zc_res['errors'][:5]: st.text(f"  {e_}")
                with st.expander("📋 每条命中详情 (真诚现货)", expanded=False):
                    for m in zc_res['matched'][:50]:
                        col = m.get('cost_col_letter', 'M')
                        st.text(f"✅ 行 {m['row']}: key={m['match_key']} "
                                f"{col} {m['old_m']} → {m['new_m']}"
                                + (f"  H '{m['old_h']}' → '{m['new_h']}'" if m.get('new_h') else ''))
        except Exception as e:
            track_cost(len(spot_items))
            st.error(f"❌ 同步真诚现货到真诚成品新单失败: {e}")
            with st.expander("详细错误"):
                st.code(traceback.format_exc())

# ---------------- 天然钻流程 (共用: 黛宝 / 布心 / 猛哥 / 二厂) ----------------
def run_natural_workflow(in_path, uploaded_name, factory_name, pt, au, gia_months=6,
                          sync_feishu=False, sync_zc_spots=True):
    """
    跑一次天然钻流程: 解析 → 查 GIA → 生成聚水潭 xlsx + 下载按钮.
    factory_name: natural.PARSERS 的键 ('黛宝'/'布心'/'猛哥'/'二厂')
    """
    st.subheader(f"💎 天然钻流程 — {factory_name}")
    parser = natural.PARSERS[factory_name]
    with st.spinner(f"解析 {factory_name} 天然钻单..."):
        try:
            items = parser(in_path, au_price=au, pt_price=pt)
            for item in items:
                item['_工厂'] = factory_name
                if factory_name == '布心':
                    item['_布心首次确认'] = buxin_initial.get(item.get('材质颜色'), False)
        except Exception as e:
            if sync_feishu: track_cost(1)
            st.error(f"❌ 解析失败: {e}")
            with st.expander("详细错误"):
                st.code(traceback.format_exc())
            return 0

    st.info(f"识别到 **{len(items)}** 件天然钻客订")
    if factory_name == '布心':
        for material in ('PT950', '18K'):
            batch_items = [i for i in items if i.get('材质颜色') == material]
            if batch_items:
                st.caption(f"{material}：{batch_items[0].get('_布心批次') or '日期待核对'} · {len(batch_items)} 件")
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
        if sync_feishu: track_cost(max(1, len(items)))
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

    with st.expander("查询明细"):
        placeholder = st.empty()
    progress = st.progress(0.0)
    lines = []
    for i, it in enumerate(items):
        gia = natural.search_gia(gia_client, order_sheets, it['款号'], it['证书号'])
        it['_gia'] = gia
        # v24.5: 把 GIA 表锁定到的真实客户名写回 item, 给 zhencheng_bitable sync 用
        # (黛宝/布心/二厂工厂单里没客户名列, 必须靠 GIA 补; 猛哥有原生客户名也覆盖也无妨)
        gia_cust = gia.get('客户名')
        if gia_cust:
            it['飞书客户名'] = gia_cust
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

    if do_jst:
        # 生成 xlsx
        jst_rows = [natural.build_jst_row(it, it['_gia'], factory_name) for it in items]
        out_path = tempfile.mktemp(suffix='.xlsx')
        date_str = datetime.now().strftime('%m-%d')
        natural.gen_jst_xlsx(jst_rows, out_path, sheet_title=f'{factory_name}真诚{date_str}')
        with open(out_path, 'rb') as f:
            data = f.read()
        try: os.unlink(out_path)
        except OSError: pass

        red_count = sum(1 for r in jst_rows if r.get('_红底'))
        if red_count:
            st.warning(f"⚠️ {red_count} 件多散货, 客户名已标红 — 请人工核验")

        fname = f'聚水潭_{factory_name}真诚_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        download_button(
            f"📥 下载 {factory_name} 真诚聚水潭 ({len(items)} 件)",
            data=data, file_name=fname,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            key=f'dl_nat_{factory_name}_{datetime.now().timestamp()}',
        )
        remember_generated("jst", factory_name + "真诚")
        st.session_state.setdefault('history', []).append({
            '时间': datetime.now().strftime('%H:%M:%S'),
            '类型': f'聚水潭-{factory_name}真诚',
            '工厂': factory_name,
            '件数': len(items),
            '文件名': fname,
            '_data': data,
        })

    # v23: 同步镶嵌成本到真诚部门多维表格 (绑定"同步今日成本到飞书"复选框)
    if sync_feishu:
        try:
            from feishu_client import FeishuClient, load_credentials
            from zhencheng_bitable import sync_zhencheng_costs

            # v28: 只同步"客户单"到真诚多维表, 现货 (E-ZC-*) 走真诚成品新单 sheet
            bitable_items = [it for it in items if it.get('类别') != '部门-真诚-现货']
            st.subheader(f"🔗 同步镶嵌成本到真诚多维表 ({len(bitable_items)} 件客户单)")
            with st.spinner("正在匹配订单并同步成本…"):
                app_id, app_secret = load_credentials()
                bitable_client = FeishuClient(app_id, app_secret)
                sync_result = sync_zhencheng_costs(bitable_client, bitable_items)

            track_bitable_cost(sync_result, len(bitable_items))
            skipped = sum(d.get('status') == 'skip' for d in sync_result['details'])
            attention = len(sync_result['not_found']) + len(sync_result['errors'])
            st.info(f"成功 {sync_result['updated']} 笔 · 跳过 {skipped} 笔 · 需要核对 {attention} 笔")
            if sync_result['not_found']:
                st.warning(f"⚠️ {len(sync_result['not_found'])} 条没匹配到: "
                           f"{', '.join(sync_result['not_found'][:10])}"
                           + ("..." if len(sync_result['not_found']) > 10 else ""))
            if sync_result['errors']:
                st.error(f"❌ {len(sync_result['errors'])} 条报错:")
                for err in sync_result['errors'][:5]:
                    st.text(f"  {err}")

            with st.expander("📋 每条详情 (含利润率)", expanded=False):
                for d in sync_result['details']:
                    ico = {'updated': '✅', 'would_update': '🔵', 'not_found': '⚠️',
                           'error': '❌', 'skip': '⏭️'}.get(d.get('status'), '·')
                    line = (f"{ico} #{d.get('no')} 证书={d.get('cert') or '空'} "
                            f"客户={d.get('name') or '空'}")
                    if d.get('matched_by'):
                        line += f" 靠[{d['matched_by']}]命中"
                    if d.get('cost') is not None:
                        line += f" 成本={d['cost']}"
                    if d.get('profit') is not None:
                        line += f" 利润={int(d['profit'])}"
                    if d.get('rate') is not None:
                        line += f" 利润率={d['rate']*100:.1f}%"
                    if d.get('note'):
                        line += d['note']
                    if d.get('reason'):
                        line += f" ({d['reason']})"
                    st.text(line)

            # v24.1: 利润简报 xlsx 下载 (只有实际写入了才生成)
            if sync_result['updated'] > 0:
                try:
                    import openpyxl
                    from openpyxl.styles import PatternFill, Font
                    from datetime import datetime as _dt
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = f'{factory_name}利润'
                    headers = ['序号', '客户名(款号)', '证书号', '品名', '材质',
                               '件数', '镶嵌成本', '飞书利润', '利润率', '备注']
                    for c, h in enumerate(headers, 1):
                        ws.cell(row=1, column=c).value = h
                        ws.cell(row=1, column=c).font = Font(bold=True)
                    RED = PatternFill('solid', fgColor='FFC7CE')
                    YEL = PatternFill('solid', fgColor='FFEB9C')
                    for i, it in enumerate(items, 2):
                        rate = it.get('飞书利润率')
                        ws.cell(row=i, column=1).value = it.get('no')
                        ws.cell(row=i, column=2).value = it.get('款号')
                        ws.cell(row=i, column=3).value = it.get('证书号')
                        ws.cell(row=i, column=4).value = it.get('品名')
                        ws.cell(row=i, column=5).value = it.get('材质颜色')
                        ws.cell(row=i, column=6).value = it.get('件数')
                        ws.cell(row=i, column=7).value = it.get('镶嵌成本')
                        ws.cell(row=i, column=8).value = it.get('飞书利润')
                        rc = ws.cell(row=i, column=9)
                        rc.value = rate
                        if rate is not None:
                            rc.number_format = '0.00%'
                            if rate < 0.15:
                                rc.fill = RED
                                ws.cell(row=i, column=10).value = '利润率偏低'
                            elif rate > 0.70:
                                rc.fill = YEL
                                ws.cell(row=i, column=10).value = '利润率异常高'
                    profit_path = tempfile.mktemp(suffix='.xlsx')
                    wb.save(profit_path)
                    with open(profit_path, 'rb') as f:
                        profit_data = f.read()
                    try: os.unlink(profit_path)
                    except OSError: pass
                    profit_fname = f'{_dt.now().strftime("%Y%m%d")}_{factory_name}_利润简报.xlsx'
                    download_button(
                        f"📊 下载利润简报 ({len(items)} 件)",
                        data=profit_data, file_name=profit_fname,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=f'dl_profit_{factory_name}_{_dt.now().timestamp()}',
                    )
                    remember_generated()
                except Exception as e:
                    st.warning(f"生成利润简报失败 (不影响主流程): {e}")

        except Exception as e:
            track_cost(len([it for it in items if it.get("类别") != "部门-真诚-现货"]))
            st.error(f"❌ 同步真诚多维表失败: {e}")
            with st.expander("详细错误"):
                st.code(traceback.format_exc())

        if sync_zc_spots:
            sync_zhencheng_finished_items(items)

    else:
        st.info(f"ℹ️ 有 {len(items)} 件可同步到真诚多维表 "
                f"(勾选「✏️ 同步今日成本到飞书」后才写入)")

    # v24.2: 工厂单完成文件 (原表追加 4 列: 总镶嵌成本/客户名/利润/利润率)
    try:
        complete_path = tempfile.mktemp(suffix='.xlsx')
        natural.write_natural_complete(in_path, items, complete_path)
        with open(complete_path, 'rb') as f:
            complete_data = f.read()
        try: os.unlink(complete_path)
        except OSError: pass
        base = uploaded_name.rsplit('.', 1)[0]
        complete_fname = f'{datetime.now().strftime("%Y%m%d")}{base}_完成.xlsx'
        download_button(
            f"📥 下载工厂单完成文件 ({factory_name}, 追加4列)",
            data=complete_data, file_name=complete_fname,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            key=f'dl_natcomplete_{factory_name}_{datetime.now().timestamp()}',
        )
        remember_generated()
        queue_archive(complete_data, in_path, items, factory_name, complete_fname, "真诚")
        st.caption("完成文件在原表最右侧追加: 总镶嵌成本 | 客户名称 | 利润 | 利润率")
        st.session_state.setdefault('history', []).append({
            '时间': datetime.now().strftime('%H:%M:%S'),
            '类型': f'工厂单完成-{factory_name}真诚',
            '工厂': factory_name,
            '件数': len(items),
            '文件名': complete_fname,
            '_data': complete_data,
        })
    except Exception as e:
        st.error(f"❌ 生成工厂单完成文件失败: {e}")
        with st.expander("详细错误"):
            st.code(traceback.format_exc())

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
    from upload_registry import CostProgress
    current_run_id = uuid.uuid4().hex
    cost_progress = CostProgress(current_run_id) if sync_feishu else None
    run_complete = False
    if cost_progress is not None:
        remember_action(cost_progress.event())
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
        elif factory_label.startswith("布心"):
            code = 'BUXIN'   # 布心走独立文件, 天然钻专属
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
        if spots_only and is_natural and code != 'E':
            st.error("天然钻单的现货不会被流水打断，不需要补写。")
            st.stop()
        if is_natural and code != 'E':
            natural_factory = detect_natural_factory_name(code, uploaded.name)
            if not natural_factory:
                st.error(f"❌ 无法定位天然钻 parser (code={code}, 文件={uploaded.name})")
                try: os.unlink(in_path)
                except OSError: pass
                st.stop()
            run_natural_workflow(in_path, uploaded.name, natural_factory, pt, au,
                                 gia_months=GIA_MONTHS, sync_feishu=sync_feishu)
            run_complete = True
            finish_run()
            try: os.unlink(in_path)
            except OSError: pass
            st.stop()

        default_material = detect_default_material(uploaded.name)
        parse_kwargs = dict(pt_price=pt, au_price=au)
        if default_material:
            parse_kwargs['default_material'] = default_material

        with st.spinner("解析工厂账单..."):
            items = factories.get_parser(code)(in_path, **parse_kwargs)

        if spots_only:
            done_before = st.session_state.get('_upload_checks_v3', {}).get(upload_key, {}).get('spot_sync')
            if done_before:
                st.warning("这份单的现货已经写过成品新单，本次没有再写。")
            else:
                sync_finished_spots(items)
            run_complete = True
            st.stop()

        clients = [it for it in items if it['类别'] == '客户单']
        spots = [it for it in items if it['类别'] in ('现货', '部门-真诚-现货')]
        repairs = [it for it in items if it['类别'] == '修理']

        st.success(f"✓ 解析 {len(items)} 件 — 客户单 {len(clients)} | 现货 {len(spots)} | 修理 {len(repairs)}")

        # E-ZC-* 由主解析器直接分流，不依赖文件名或序号栏写“19楼”。
        zc_spots = [it for it in items if it['类别'] == '部门-真诚-现货']
        if zc_spots:
            if sync_feishu:
                sync_zhencheng_finished_items(zc_spots)
            else:
                st.info(f"ℹ️ {len(zc_spots)} 件真诚现货 → 真诚成品新单 M 列（未勾选同步，不写入）")

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
                # v20.6: "退货中" 也视为退货 (跟"已退款"同处理)
                is_refunded = any(kw in status_text for kw in ('已退款', '退货中', '退货'))
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
                            final_cost = math.ceil(existing_cost + today_cost)
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

        if sync_feishu:
            track_cost(len(clients), write_ok)

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
                download_button(
                    "📥 下载聚水潭入库 Excel", data=jst_data, file_name=jst_fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type="primary",
                    key=f'dl_jst_{datetime.now().timestamp()}',
                )
                remember_generated("jst", code + "常规")
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
            download_button(
                "📥 下载工厂单 _完成.xlsx",
                data=done_data, file_name=done_fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f'dl_done_{datetime.now().timestamp()}',
                type="primary",
            )
            remember_generated()
            queue_archive(done_data, in_path, items, code, done_fname, "常规")
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
                                     gia_months=GIA_MONTHS, sync_feishu=sync_feishu, sync_zc_spots=False)
            except Exception as e:
                if sync_feishu: track_cost(1)
                st.error(f"❌ 天然钻流程失败: {e}")
                with st.expander("详细错误"):
                    st.code(traceback.format_exc())

        # v25: 培育钻文件里 类别='部门-真诚' 的件 → 真诚多维表 (跟天然钻真诚一起入19楼珠宝销售)
        # 猛哥 A 列 = '19楼' 的行, factories.parse_E 会标 类别='部门-真诚'; B 列 = 真实客户名
        if sync_feishu:
            zc_pei_items = []
            for it in items:
                if it.get('类别') != '部门-真诚':
                    continue
                it2 = dict(it)
                if code == 'E':
                    it2['_工厂'] = '猛哥'
                # factories 硬编码 飞书客户名='真诚' (假名), 用 B 列 单号 = 真实客户名覆盖
                b_col = str(it.get('单号') or '').strip()
                if b_col and b_col != '19楼':
                    it2['飞书客户名'] = b_col
                else:
                    it2['飞书客户名'] = None   # 让 sync 走证书号
                zc_pei_items.append(it2)

            if zc_pei_items:
                try:
                    from zhencheng_bitable import sync_zhencheng_costs
                    st.divider()
                    st.subheader(f"🔗 真诚(19楼)培育件 → 真诚多维表 ({len(zc_pei_items)} 件)")
                    with st.spinner("同步真诚培育件..."):
                        zc_res = sync_zhencheng_costs(_client, zc_pei_items)
                    track_bitable_cost(zc_res, len(zc_pei_items))
                    st.success(
                        f"✅ 匹配 **{zc_res['matched']}** / 写入 **{zc_res['updated']}** "
                        f"| 证书=`{zc_res.get('cert_field') or '(无)'}` "
                        f"客户名=`{zc_res.get('name_field') or '(无)'}` "
                        f"成本=`{zc_res['cost_field']}`"
                    )
                    if zc_res['not_found']:
                        st.warning(f"⚠️ {len(zc_res['not_found'])} 件没匹配到: "
                                   f"{', '.join(zc_res['not_found'][:8])}")
                    if zc_res['errors']:
                        st.error(f"❌ {len(zc_res['errors'])} 件错误")
                        for e in zc_res['errors'][:5]: st.text(f"  {e}")
                    with st.expander("📋 每条详情 (真诚培育)", expanded=False):
                        for d in zc_res['details']:
                            ico = {'updated':'✅','not_found':'⚠️','error':'❌','skip':'⏭️'}.get(d.get('status'),'·')
                            line = f"{ico} #{d.get('no')} 客户={d.get('name') or '空'}"
                            if d.get('matched_by'): line += f" 靠[{d['matched_by']}]"
                            if d.get('cost') is not None: line += f" 成本={d['cost']}"
                            if d.get('profit') is not None: line += f" 利润={int(d['profit'])}"
                            if d.get('rate') is not None: line += f" 利润率={d['rate']*100:.1f}%"
                            if d.get('note'): line += d['note']
                            if d.get('reason'): line += f" ({d['reason']})"
                            st.text(line)
                except Exception as e:
                    track_cost(len(zc_pei_items))
                    st.error(f"❌ 同步真诚培育件失败: {e}")
                    with st.expander("详细错误"):
                        st.code(traceback.format_exc())

        # v20 + v22.9: 同步镶嵌成本到成品新单 (飞书电子表格) M 列
        # v22.9: 绑定 "✏️同步今日成本到飞书" 复选框, 不勾就不跑 (避免重复叠加)
        if sync_feishu:
            sync_finished_spots(items)
        else:
            # v22.9: 不勾同步飞书 → 只提示不动
            sync_items_preview = None
            try:
                from finished_products import build_sync_items_from_factory_items
                sync_items_preview = build_sync_items_from_factory_items(items)
            except Exception:
                pass
            if sync_items_preview:
                st.divider()
                st.info(f"ℹ️ 有 {len(sync_items_preview)} 条现货可同步到成品新单 (勾选「✏️ 同步今日成本到飞书」后才写入)")

        run_complete = True
        finish_run()
        try:
            os.unlink(in_path)
        except OSError:
            pass

    except Exception as e:
        st.error(f"❌ 出错: {e}")
        with st.expander("详细错误"):
            st.code(traceback.format_exc())
    finally:
        if cost_progress is not None:
            remember_action(cost_progress.event(complete=run_complete))

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
            download_button(
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

**.xls 文件**: 可直接上传，网站会自动转换

**✏️ 同步飞书镶嵌成本** (v12):
- 默认**不勾**, 只读飞书不写
- 勾上后会把今天工厂账单算的成本写到飞书「镶嵌成本」字段
- 默认**叠加** (原值 + 今天值), 用于同一单分多次出货的场景
- 想直接替换原值就勾「覆盖模式」
- 飞书"货物状态"含"已退款"的件 → 自动写 0 + 客户名加"已退款做现货"
    """)
