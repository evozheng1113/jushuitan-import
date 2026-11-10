"""工厂流水：把本批完成单原样导入成飞书表格，再在「流水总表」往下记一行。
比逐格写格式快且稳：格式、合并、图片由飞书导入一次完成。"""
import json
import time
from datetime import datetime, timezone, timedelta
import requests
from factory_archive import TOKEN, ArchiveAPIError, ArchiveUncertain, checked

BASE = "https://open.feishu.cn/open-apis"
INDEX_TITLE = "流水总表"
INDEX_HEADER = ["导入时间", "工厂", "出货单", "件数", "链接", "批次"]
FOLDER_PREFIX = "工厂流水-"


class FlowImporter:
    def __init__(self, client):
        self.client = client
        self.session = requests.Session()  # 复用连接：服务器在海外，每次重新握手很慢

    def call(self, method, path, **kwargs):
        headers = self.client._headers(json_body="json" in kwargs)
        return checked(self.session.request(method, BASE + path, headers=headers, timeout=60, **kwargs))

    def folder(self, factory):
        """应用自己云空间里的「工厂流水-工厂名」文件夹；找不到就建，出错就放根目录。"""
        name = FOLDER_PREFIX + factory
        try:
            root = self.call("GET", "/drive/explorer/v2/root_folder/meta")["token"]
            page = None
            while True:
                params = {"folder_token": root, "page_size": 200}
                if page:
                    params["page_token"] = page
                data = self.call("GET", "/drive/v1/files", params=params)
                for item in data.get("files") or []:
                    if item.get("type") == "folder" and item.get("name") == name:
                        return item["token"]
                if not data.get("has_more"):
                    break
                page = data.get("next_page_token")
            return self.call("POST", "/drive/v1/files/create_folder",
                             json={"name": name, "folder_token": root})["token"]
        except (ArchiveAPIError, requests.exceptions.RequestException, KeyError):
            return ""

    def upload(self, data, name):
        extra = json.dumps({"obj_type": "sheet", "file_extension": "xlsx"})
        result = self.call("POST", "/drive/v1/medias/upload_all",
                           data={"file_name": name, "parent_type": "ccm_import_open",
                                 "size": str(len(data)), "extra": extra},
                           files={"file": (name, data)})
        return result["file_token"]

    def run_import(self, file_token, title, folder, wait=120):
        ticket = self.call("POST", "/drive/v1/import_tasks", json={
            "file_extension": "xlsx", "file_token": file_token, "type": "sheet", "file_name": title,
            "point": {"mount_type": 1, "mount_key": folder}})["ticket"]
        deadline = time.time() + wait
        while True:
            try:
                result = self.call("GET", f"/drive/v1/import_tasks/{ticket}")["result"]
            except (ArchiveAPIError, requests.exceptions.RequestException):
                result = {"job_status": 2}
            status = result.get("job_status")
            if status == 0:
                return result["token"], result["url"]
            if status not in (1, 2):
                raise RuntimeError(f"飞书导入失败：{result.get('job_error_msg') or status}")
            if time.time() > deadline:
                raise ArchiveUncertain("飞书导入超时，可能稍后才出现；请先查看工厂流水文件夹，勿直接重跑")
            time.sleep(1.5)

    def share(self, token):
        """组织内获得链接的人可阅读；失败只影响能否点开，不影响已导入。"""
        try:
            self.call("PATCH", f"/drive/v1/permissions/{token}/public", params={"type": "sheet"},
                      json={"link_share_entity": "tenant_readable"})
            return True
        except (ArchiveAPIError, requests.exceptions.RequestException):
            return False

    def index(self, row):
        """总表每行 A 列都有时间，飞书追加接口会正确接在最后一行下面。"""
        sheets = self.call("GET", f"/sheets/v3/spreadsheets/{TOKEN}/sheets/query").get("sheets", [])
        match = [s for s in sheets if s.get("title") == INDEX_TITLE]
        if not match:
            self.call("POST", f"/sheets/v2/spreadsheets/{TOKEN}/sheets_batch_update",
                      json={"requests": [{"addSheet": {"properties": {"title": INDEX_TITLE}}}]})
            sheets = self.call("GET", f"/sheets/v3/spreadsheets/{TOKEN}/sheets/query").get("sheets", [])
            match = [s for s in sheets if s.get("title") == INDEX_TITLE]
            sid = match[0]["sheet_id"]
            values = [INDEX_HEADER, row]
        else:
            sid = match[0]["sheet_id"]
            values = [row]
        self.call("POST", f"/sheets/v2/spreadsheets/{TOKEN}/values_append",
                  params={"insertDataOption": "INSERT_ROWS"},
                  json={"valueRange": {"range": f"{sid}!A1:F1", "values": values}})


def import_batch(client, book, factory, filename, count, batch, progress=None):
    """返回 (链接, 提示列表)。导入成功后的步骤失败只提示，不当作失败。"""
    steps = [0]

    def tick(total=5):
        steps[0] += 1
        if progress:
            progress(min(steps[0], total), total)

    importer = FlowImporter(client)
    now = datetime.now(timezone(timedelta(hours=8)))
    base = filename.rsplit(".", 1)[0]
    title = f"{factory}-{now:%Y%m%d-%H%M}-{base}"[:120]
    folder = importer.folder(factory); tick()
    file_token = importer.upload(book, title + ".xlsx"); tick()
    token, url = importer.run_import(file_token, title, folder); tick()
    notes = []
    if not importer.share(token):
        notes.append("链接分享没设置成功，同事可能打不开，请在飞书里手动分享。")
    tick()
    try:
        importer.index([now.strftime("%Y-%m-%d %H:%M"), factory, filename, count, url, batch[:12]])
    except Exception as exc:
        notes.append(f"流水总表没有记上这一行（{str(exc)[:80]}），表格已导入：{url}")
    tick()
    return url, notes
