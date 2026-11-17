"""把工厂单完成文件发到飞书群（机器人就是网站的飞书应用「工厂出货镶嵌成本自动化」，用户 2026-11-17 定）。

- 用网站同一个飞书应用发，应用要开机器人能力和发消息、上传文件权限，机器人要先拉进群。
- 只发到群「工厂单汇总 查询」（用户 2026-11-17 建），比较群名时忽略空格；
  secrets 里填 FEISHU_GROUP_NAME 可以换群。机器人被拉进别的群也不会发过去。
- 只发文件，不写飞书表格；发送失败只提示，不影响成本、完成单和流水。
"""
import json
import os

import requests

BASE = "https://open.feishu.cn/open-apis"
MAX_BYTES = 30 * 1024 * 1024  # 飞书文件消息上限


class GroupError(RuntimeError):
    pass


def _checked(response):
    try:
        payload = response.json()
    except ValueError:
        raise GroupError(f"飞书返回异常（HTTP {getattr(response, 'status_code', '?')}）")
    if payload.get("code") != 0:
        raise GroupError(f"{payload.get('code')} {payload.get('msg', '')}".strip())
    return payload.get("data") or {}


def _secret(key):
    value = os.environ.get(key)
    try:
        import streamlit as st
        value = st.secrets.get(key) or value
    except Exception:
        pass
    return str(value).strip() if value else None


def bot_credentials():
    """机器人如果是另建的飞书应用，在网站后台填 FEISHU_BOT_APP_ID / FEISHU_BOT_APP_SECRET；没填就用网站的应用。"""
    app_id, secret = _secret("FEISHU_BOT_APP_ID"), _secret("FEISHU_BOT_APP_SECRET")
    return (app_id, secret) if app_id and secret else None


DEFAULT_GROUP = "工厂单汇总 查询"


def _group_name():
    return _secret("FEISHU_GROUP_NAME") or DEFAULT_GROUP


def _plain(name):
    return "".join(str(name or "").split())


def find_chat(headers, name=None):
    """返回 (chat_id, 群名)。只在机器人所在的群里找。"""
    chats, token = [], None
    while True:
        params = {"page_size": 100}
        if token:
            params["page_token"] = token
        data = _checked(requests.get(f"{BASE}/im/v1/chats", headers=headers, params=params, timeout=20))
        chats += [(c.get("chat_id"), c.get("name") or "") for c in data.get("items") or []]
        token = data.get("page_token")
        if not data.get("has_more") or not token:
            break
    if name:
        hits = [c for c in chats if _plain(c[1]) == _plain(name)]
        if len(hits) != 1:
            raise GroupError(f"机器人{'还没被拉进' if not hits else '在多个叫'}「{name}」{'群' if not hits else '的群里'}")
        return hits[0]
    if not chats:
        raise GroupError("机器人还没被拉进任何群")
    if len(chats) > 1:
        raise GroupError("机器人在多个群里，请在网站后台设置 FEISHU_GROUP_NAME 指定发到哪个群")
    return chats[0]


def send_file(headers, data, filename, chat=None):
    """上传文件并发到群，返回群名。headers 只需带 Authorization。"""
    if len(data) > MAX_BYTES:
        raise GroupError("文件超过 30MB，飞书不让发")
    auth = {"Authorization": headers["Authorization"]}
    chat_id, chat_name = chat or find_chat(auth, _group_name())
    up = _checked(requests.post(f"{BASE}/im/v1/files", headers=auth, timeout=60,
                                data={"file_type": "xls", "file_name": filename},
                                files={"file": (filename, data)}))
    if not up.get("file_key"):
        raise GroupError("上传文件没有拿到 file_key")
    _checked(requests.post(f"{BASE}/im/v1/messages", params={"receive_id_type": "chat_id"},
                           headers={**auth, "Content-Type": "application/json"}, timeout=30,
                           json={"receive_id": chat_id, "msg_type": "file",
                                 "content": json.dumps({"file_key": up["file_key"]})}))
    return chat_name
