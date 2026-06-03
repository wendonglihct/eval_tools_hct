import lark_oapi as lark
from lark_oapi.api.im.v1 import *
from lark_oapi.api.drive.v1 import *
from lark_oapi.api.wiki.v2 import *
from lark_oapi.api.contact.v3 import *
import json
import os
import re
import ast
import time
import requests
import sys
from requests_toolbelt import MultipartEncoder

PROJECT_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..")
)
sys.path.insert(0, PROJECT_ROOT)
from report.hct_report import HCTReportBuilder
from report.config_loader import load_mode_routing
from report.debug_report import process_debug_report
from concurrent.futures import ThreadPoolExecutor
from utils.log_utils import setup_bot_logger
from utils.paths import LOG_DIR, config_path

logger = setup_bot_logger(LOG_DIR)
logger.info("bot started")
# 飞书开放平台应用的 App ID 和 App Secret
APP_ID = "cli_a91545943de31bd9"
APP_SECRET = "qLUjkfbfisTW2eejmur6UgIDObjddo1t"

# 全局线程池
executor = ThreadPoolExecutor(max_workers=20)

# 事件去重窗口（飞书重试会复用 event_id，用 set 拦截）
memory = {"events_id": set()}


class MessageParser:
    """飞书消息文本 → (tasks_dict, ver_map, mode) 解析，以及事件时间窗去重判断。

    纯函数集合，无状态；放成类只是为了让 robot.py 更结构化、便于后续扩展（如多种 DSL 语法）。
    """

    # 200 秒之前的事件不处理（飞书重试也走这里，避免历史事件被重复处理）
    EVENT_TIME_WINDOW_MS = 200_000

    @staticmethod
    def is_event_stale(event_create_time_ms) -> bool:
        current_ms = int(time.time() * 1e3)
        return abs(int(event_create_time_ms) - current_ms) >= MessageParser.EVENT_TIME_WINDOW_MS

    @staticmethod
    def parse(res_content: str):
        """提取 tasks_dict / ver_map / mode（mode 自动大写，缺省 DT_HCT）。"""
        tasks_dict_str = re.search(r'tasks_dict\s*=\s*(\{.*?\})', res_content, re.S).group(1)
        tasks_dict = ast.literal_eval(tasks_dict_str)

        ver_map_match = re.search(r'ver_map\s*=\s*(\[.*?\])', res_content, re.S)
        ver_map = ast.literal_eval(ver_map_match.group(1)) if ver_map_match else []

        mode_match = re.search(r'mode\s*=\s*["\']?(.*?)["\']?\s*$', res_content, re.M)
        mode = mode_match.group(1).upper() if mode_match else "DT_HCT"
        return tasks_dict, ver_map, mode


def _check_lark(response, action):
    """飞书 OpenAPI 返回的统一成功检查：失败时按既定格式记 error 日志并返回 False。"""
    if response.success():
        return True
    try:
        body = json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)
    except Exception:
        body = str(getattr(response.raw, "content", ""))
    lark.logger.error(
        f"{action} failed, code: {response.code}, msg: {response.msg}, "
        f"log_id: {response.get_log_id()}, resp: \n{body}"
    )
    return False


class FeishuClient:
    """飞书 OpenAPI 调用聚合。集中所有 token / drive / wiki / im / contact 操作；
    业务侧通过单例 `feishu` 调用即可，避免全局 lark client 散落各处。

    线程安全：内部不持任何可变态，仅持外部 lark.Client；所有方法可重入。
    """

    # 飞书云空间挂载目标 / wiki 空间常量
    PARENT_NODE = "Col0fBF8slvQrFdzEfAcAWygnYs"
    WIKI_SPACE_ID = "7560534433279868930"
    WIKI_PARENT_TOKEN = "GOgIwx6N8iutB5kIlXGcHuGEnnh"
    WIKI_URL_PREFIX = "https://horizonrobotics.feishu.cn/wiki/"

    def __init__(self, app_id, app_secret, lark_client):
        self.app_id = app_id
        self.app_secret = app_secret
        self._client = lark_client

    # ---------- auth ----------
    def get_tenant_access_token(self):
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        payload = {"app_id": self.app_id, "app_secret": self.app_secret}
        try:
            resp = requests.post(url, json=payload, timeout=5)
            resp.raise_for_status()
        except requests.RequestException as e:
            raise RuntimeError(f"请求飞书接口失败: {e}")
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"获取 tenant_access_token 失败: {data}")
        return data["tenant_access_token"], data["expire"]

    # ---------- drive ----------
    def upload_file(self, file_path, file_name, te_token):
        file_size = os.path.getsize(file_path)
        url = "https://open.feishu.cn/open-apis/drive/v1/files/upload_all"
        with open(file_path, "rb") as fp:
            form = {
                "file_name": file_name,
                "parent_type": "explorer",
                "parent_node": self.PARENT_NODE,
                "size": str(file_size),
                "file": fp,
            }
            multi_form = MultipartEncoder(form)
            headers = {"Authorization": f"Bearer {te_token}", "Content-Type": multi_form.content_type}
            response = requests.request("POST", url, headers=headers, data=multi_form)
        time.sleep(5)
        return json.loads(response.text)["data"]["file_token"]

    def transfer_xlsx_to_sheet(self, file_token, file_name):
        req = (
            CreateImportTaskRequest.builder()
            .request_body(
                ImportTask.builder()
                .file_extension("xlsx")
                .file_token(file_token)
                .type("sheet")
                .file_name(file_name)
                .point(ImportTaskMountPoint.builder().mount_type(1).mount_key(self.PARENT_NODE).build())
                .build()
            )
            .build()
        )
        resp: CreateImportTaskResponse = self._client.drive.v1.import_task.create(req)
        if not _check_lark(resp, "client.drive.v1.import_task.create"):
            return
        time.sleep(5)
        req = GetImportTaskRequest.builder().ticket(resp.data.ticket).build()
        resp: GetImportTaskResponse = self._client.drive.v1.import_task.get(req)
        if not _check_lark(resp, "client.drive.v1.import_task.get"):
            return
        return resp.data.result.token

    # ---------- wiki ----------
    def move_to_wiki(self, file_token, file_type):
        req = (
            MoveDocsToWikiSpaceNodeRequest.builder()
            .space_id(self.WIKI_SPACE_ID)
            .request_body(
                MoveDocsToWikiSpaceNodeRequestBody.builder()
                .parent_wiki_token(self.WIKI_PARENT_TOKEN)
                .obj_type(file_type)
                .obj_token(file_token)
                .build()
            )
            .build()
        )
        resp: MoveDocsToWikiSpaceNodeResponse = self._client.wiki.v2.space_node.move_docs_to_wiki(req)
        if not _check_lark(resp, "client.wiki.v2.space_node.move_docs_to_wiki"):
            return
        time.sleep(5)
        task_id = resp.data.task_id
        req = GetTaskRequest.builder().task_id(task_id).task_type("move").build()
        resp: GetTaskResponse = self._client.wiki.v2.task.get(req)
        if not _check_lark(resp, "client.wiki.v2.task.get"):
            return
        return f"{self.WIKI_URL_PREFIX}{resp.data.task.move_result[0].node.node_token}"

    # ---------- contact ----------
    def get_user_name(self, open_id):
        if not open_id:
            return "unknown"
        try:
            req = GetUserRequest.builder().user_id(open_id).user_id_type("open_id").build()
            resp = self._client.contact.v3.user.get(req)
            if not resp.success():
                logger.warning(
                    f"get user name failed: code={resp.code}, msg={resp.msg}, "
                    f"log_id={resp.get_log_id()}"
                )
                return "unknown"
            return getattr(resp.data.user, "name", None) or "unknown"
        except Exception as e:
            logger.exception(f"get_user_name 异常: {e}")
            return "unknown"

    # ---------- im ----------
    def _create_or_reply(self, data, content, msg_type):
        """根据 chat_type 自动走 create（p2p）或 reply（群聊）。"""
        if data.event.message.chat_type == "p2p":
            req = (
                CreateMessageRequest.builder()
                .receive_id_type("chat_id")
                .request_body(
                    CreateMessageRequestBody.builder()
                    .receive_id(data.event.message.chat_id)
                    .uuid(data.header.event_id)
                    .msg_type(msg_type)
                    .content(content)
                    .build()
                )
                .build()
            )
            return self._client.im.v1.message.create(req)
        req = (
            ReplyMessageRequest.builder()
            .message_id(data.event.message.message_id)
            .request_body(
                ReplyMessageRequestBody.builder().content(content).msg_type(msg_type).build()
            )
            .build()
        )
        return self._client.im.v1.message.reply(req)

    def send_text(self, data, text):
        return self._create_or_reply(data, json.dumps({"text": text}), "text")

    def send_card(self, data, content):
        return self._create_or_reply(data, content, "interactive")


class CardBuilder:
    """交互卡片渲染。当前仅 DT_HCT 用 card_sv.json，未来不同 mode 可用不同模板。"""

    @classmethod
    def _load_template(cls, name: str) -> dict:
        with open(config_path(name), "r", encoding="utf-8") as f:
            return json.load(f)

    @classmethod
    def hct_result_card(cls, sheet_url: str) -> str:
        """注入 sheet 链接到 card_sv.json 的第一个 text 元素，返回 JSON 字符串。"""
        data = cls._load_template("card_sv.json")
        data["body"]["elements"][0]["content"] = f"[评测结果-表格]({sheet_url})"
        return json.dumps(data, ensure_ascii=False)


class UserLookup:
    """open_id → 用户姓名 / 安全文件名片段。失败时返回 fallback，永不抛异常给上游。"""

    # 文件名安全字符正则：保留中英文/数字/下划线/点；其他统一替换为 _
    _FILENAME_UNSAFE_RE = re.compile(r"[^\w.\u4e00-\u9fff]+")

    def __init__(self, feishu_client):
        self._feishu = feishu_client

    def name(self, open_id: str) -> str:
        return self._feishu.get_user_name(open_id)

    @classmethod
    def safe_filename(cls, text: str, fallback: str = "unknown") -> str:
        t = cls._FILENAME_UNSAFE_RE.sub("_", str(text or "")).strip("_")
        return t if t else fallback


def _run_dt_hct(data: P2ImMessageReceiveV1, tasks_dict, ver_map, mode):
    """DT_HCT 主流程：下载 → 对比 → 上传 → 转 sheet → 入 wiki → 卡片回复。"""
    te_token, _ = feishu.get_tenant_access_token()

    file_name_suffix = time.strftime("%Y%m%d%H%M%S", time.localtime())
    xlsx_path = HCTReportBuilder.build_report(tasks_dict, ver_map, file_name_suffix, mode)

    # 取发送者 open_id → 用户名 → 拼入文件名
    sender = getattr(data.event, "sender", None)
    sender_id = getattr(sender, "sender_id", None) if sender else None
    open_id = getattr(sender_id, "open_id", None) if sender_id else None
    user_name = UserLookup.safe_filename(feishu.get_user_name(open_id))
    feishu_file_name = f"多版本评测结果_{mode}_{user_name}_{file_name_suffix}.xlsx"
    logger.info(f"飞书文件名: {feishu_file_name}  (user={user_name}, open_id={open_id})")

    xlsx_token = feishu.upload_file(xlsx_path, feishu_file_name, te_token)
    feish_sheet_token = feishu.transfer_xlsx_to_sheet(xlsx_token, feishu_file_name)
    sheet_url = feishu.move_to_wiki(feish_sheet_token, "sheet")

    # 发送原卡片（wiki 链接）
    response = feishu.send_card(data, CardBuilder.hct_result_card(sheet_url))
    if not response.success():
        raise Exception(f"消息发送失败: {response.code}, {response.msg}, log_id: {response.get_log_id()}")

    # 发送简表卡片
    # card_json = process_debug_report(xlsx_path, sheet_url)
    # card_content = json.dumps(card_json, ensure_ascii=False)
    # response = feishu.send_card(data, card_content)
    # if not response.success():
    #     raise Exception(f"简表卡片发送失败: {response.code}, {response.msg}, log_id: {response.get_log_id()}")

class Dispatcher:
    """根据 modes.json 的 handler_groups 把 mode 派发到对应 handler。

    handler 注册采用类属性字典 HANDLERS，便于后续按 mode 注册新业务而无需改本类。
    对外暴露 dispatch(data, res_content) 一个方法即可：
      - 解析消息
      - 查 handler_key
      - 三态：handler 已注册 → 执行；handler_key 找到但未注册 → 「功能开发中」；无 handler_key → 「mode 错误」
      - 任意异常 → 回复 "评测对比任务失败"
    """

    HANDLERS = {}  # handler_key -> callable(data, tasks_dict, ver_map, mode)

    def __init__(self, feishu_client):
        self._feishu = feishu_client

    @classmethod
    def register(cls, handler_key: str, fn):
        cls.HANDLERS[handler_key] = fn

    def dispatch(self, data, res_content: str):
        try:
            logger.info(f"输入内容：{res_content}")
            tasks_dict, ver_map, mode = MessageParser.parse(res_content)

            routing = load_mode_routing()
            groups = routing.get("handler_groups", {}) or {}

            handler_key = next(
                (k for k, modes in groups.items() if mode in (modes or [])),
                None,
            )

            if handler_key is None:
                all_supported = sorted({m for modes in groups.values() for m in (modes or [])})
                logger.warning(f"mode={mode} 不在任何 handler_group 中")
                self._feishu.send_text(
                    data,
                    f"mode 输入错误，暂不支持该值：{mode}\n当前已支持：{all_supported}"
                )
                return

            handler = self.HANDLERS.get(handler_key)
            if handler is None:
                logger.info(f"mode={mode} 命中 handler_key={handler_key}，但 handler 未实现")
                self._feishu.send_text(
                    data, f"功能开发中，暂未上线：mode={mode}（handler={handler_key}）"
                )
                return

            handler(data, tasks_dict, ver_map, mode)
        except Exception as e:
            logger.exception("处理消息失败")
            try:
                self._feishu.send_text(data, f"评测对比任务失败：{e}")
            except Exception:
                logger.exception("回复异常消息失败")


def do_p2_im_message_receive_v1(data: P2ImMessageReceiveV1) -> None:
    """
    主回调函数：快速响应，将任务丢到线程池
    """
    res_content = ""
    if data.event.message.message_type == "text":
        res_content = json.loads(data.event.message.content)["text"]
    else:
        res_content = "解析消息失败，请发送文本消息\nparse message failed, please send text message"
    # 去重
    if data.header.event_id in memory["events_id"] or MessageParser.is_event_stale(data.header.create_time):
        return
    else:
        memory["events_id"].add(data.header.event_id)
    logger.info(f"Received a new message, start processing event_id={data.header.event_id}")
    executor.submit(dispatcher.dispatch, data, res_content)


# 注册事件回调
event_handler = (
    lark.EventDispatcherHandler.builder("", "")
    .register_p2_im_message_receive_v1(do_p2_im_message_receive_v1)
    .build()
)


# 创建 LarkClient 和 LarkWSClient
client = lark.Client.builder().app_id(APP_ID).app_secret(APP_SECRET).build()
# 实例化 FeishuClient 单例（旧的模块级函数全部通过它代理）
feishu = FeishuClient(APP_ID, APP_SECRET, client)
# 实例化 Dispatcher 单例并注册 handler（handler_key 来自 modes.json/handler_groups）
dispatcher = Dispatcher(feishu)
Dispatcher.register("DT_HCT", _run_dt_hct)
# 后续: Dispatcher.register("DT_HCT_CLOSE", _run_dt_hct_close)
wsClient = lark.ws.Client(
    APP_ID,
    APP_SECRET,
    event_handler=event_handler,
    log_level=lark.LogLevel.DEBUG,
)


def main():
    wsClient.start()


if __name__ == "__main__":
    main()
