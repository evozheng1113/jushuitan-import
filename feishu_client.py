"""飞书 API 客户端 — 网站版
凭证来源 (优先级):
1. Streamlit Cloud Secrets (生产环境, 推荐)
2. 环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET
3. 本地 feishu_secret.docx (开发机)

⚠️ 安全提醒
- 永远不要把 APP_ID / APP_SECRET 写在代码里 / push 到 GitHub
- Streamlit Cloud 后台 Secrets 是加密存的, 仓库 public 也安全
- 你的 Streamlit 账号必须开两步验证
"""
import os
import time
import requests


# 飞书 Base 表定位 (这两个 ID 不算敏感, 单独看没用)
APP_TOKEN = 'LvnrbfG64areQPsccHYcXmf7njd'   # 珠宝销售副本
TABLE_ID = 'tbl2GcSHhDqgQzv7'               # 收支表


def load_credentials():
    """按优先级读凭证, 找不到时抛清晰的错误"""
    # 1. Streamlit Cloud Secrets
    try:
        import streamlit as st
        if hasattr(st, 'secrets'):
            try:
                app_id = st.secrets.get('FEISHU_APP_ID') or st.secrets.get('APP_ID')
                app_secret = st.secrets.get('FEISHU_APP_SECRET') or st.secrets.get('APP_SECRET')
                if app_id and app_secret:
                    return str(app_id), str(app_secret)
            except (FileNotFoundError, KeyError):
                pass
    except ImportError:
        pass

    # 2. 环境变量
    if os.environ.get('FEISHU_APP_ID') and os.environ.get('FEISHU_APP_SECRET'):
        return os.environ['FEISHU_APP_ID'], os.environ['FEISHU_APP_SECRET']

    # 3. 本地 docx 兜底
    secret_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'feishu_secret.docx')
    if os.path.exists(secret_path):
        from docx import Document
        doc = Document(secret_path)
        app_id = app_secret = ''
        for p in doc.paragraphs:
            t = p.text.strip()
            if t.startswith('APP_ID='):
                app_id = t.split('=', 1)[1].strip()
            elif t.startswith('APP_SECRET='):
                app_secret = t.split('=', 1)[1].strip()
        if app_id and app_secret:
            return app_id, app_secret

    raise RuntimeError(
        "飞书凭证未配置。\n"
        "1) Streamlit Cloud: 在 App 的 Settings → Secrets 里加 "
        "FEISHU_APP_ID='cli_xxx' / FEISHU_APP_SECRET='xxx'\n"
        "2) 本地开发: 设置环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET, "
        "或者在脚本同目录放 feishu_secret.docx"
    )


class FeishuClient:
    def __init__(self, app_id, app_secret):
        self.app_id = app_id
        self.app_secret = app_secret
        self._token = None
        self._token_expire = 0

    def _get_token(self):
        if self._token and time.time() < self._token_expire:
            return self._token
        r = requests.post(
            'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
            json={'app_id': self.app_id, 'app_secret': self.app_secret},
            timeout=10,
        )
        data = r.json()
        if data.get('code') != 0:
            raise RuntimeError(f"飞书 token 获取失败: {data}")
        self._token = data['tenant_access_token']
        self._token_expire = time.time() + data['expire'] - 60
        return self._token

    def _headers(self):
        return {'Authorization': f'Bearer {self._get_token()}',
                'Content-Type': 'application/json'}

    # ============ 通用工具 ============
    @staticmethod
    def get_text(v):
        """从飞书字段值里提取纯文本"""
        if v is None:
            return None
        if isinstance(v, list) and v:
            first = v[0]
            if isinstance(first, dict):
                return first.get('text', '') or first.get('name', '') or ''
            return str(first)
        if isinstance(v, dict):
            return v.get('text', '') or v.get('name', '') or str(v)
        return str(v)

    @staticmethod
    def get_number(v):
        """从飞书字段值里提取数字, 失败返回 None
        v14.4: 支持嵌套 list / dict (公式字段返回 [{type:'number', text:'1393', value:1393}])
        """
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return v
        if isinstance(v, str):
            try:
                return float(v.replace(',', '').replace('¥', '').strip())
            except ValueError:
                return None
        if isinstance(v, list) and v:
            return FeishuClient.get_number(v[0])
        if isinstance(v, dict):
            # 飞书公式字段常见结构: {value: [...], type: 'Formula'}
            #                   或 {value: 1393, type: 'Number'}
            #                   或 {text: '1393', value: 1393}
            for key in ('value', 'number', 'text', 'name'):
                if key in v:
                    n = FeishuClient.get_number(v[key])
                    if n is not None:
                        return n
        return None

    # ============ 查询 ============
    def find_by_field(self, app_token, table_id, field_name, value):
        url = (f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}'
               f'/tables/{table_id}/records/search')
        body = {
            "filter": {
                "conjunction": "and",
                "conditions": [{
                    "field_name": field_name,
                    "operator": "is",
                    "value": [value],
                }],
            }
        }
        r = requests.post(url, headers=self._headers(), json=body, timeout=10)
        items = r.json().get('data', {}).get('items', [])
        return items[0] if items else None

    def find_by_order_number(self, app_token, table_id, order_number):
        return self.find_by_field(app_token, table_id, '下单单号', order_number)

    def find_by_cert(self, app_token, table_id, cert):
        # 飞书"证书编码"字段, 兜底用
        for fname in ('证书编码', '证书编号'):
            try:
                rec = self.find_by_field(app_token, table_id, fname, cert)
                if rec:
                    return rec
            except Exception:
                continue
        return None

    # ============ 写入 (网站默认不调, 但保留) ============
    def update_record(self, app_token, table_id, record_id, fields):
        url = (f'https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}'
               f'/tables/{table_id}/records/{record_id}')
        r = requests.put(url, headers=self._headers(),
                         json={"fields": fields}, timeout=10)
        return r.json()
