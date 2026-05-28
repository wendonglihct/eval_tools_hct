import lark_oapi as lark
from lark_oapi.api.im.v1 import *
from lark_oapi.api.drive.v1 import *
from lark_oapi.api.wiki.v2 import *
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
from report.hct_report import run_compare_hct
from concurrent.futures import ThreadPoolExecutor
from utils.log_utils import setup_bot_logger
from threading import Lock

tasks_lock = Lock()
logger = setup_bot_logger(os.path.join(PROJECT_ROOT, "robot_memory/log"))
logger.info("bot started")
# 飞书开放平台应用的 App ID 和 App Secret
APP_ID = "cli_a91545943de31bd9"
APP_SECRET = "qLUjkfbfisTW2eejmur6UgIDObjddo1t"

# 全局线程池
executor = ThreadPoolExecutor(max_workers=20)

memory = {
    "events_id": set(),   # 用 set 查询更快
    "chats": {}           # 存对话
}


def time_diff(event_time):
    current_time_seconds = time.time()
    current_time_milliseconds = int(current_time_seconds * 1e3)
    # 200s之前的事件不处理
    if abs(int(event_time) - current_time_milliseconds) >= 200000:
        return True
    return False


def parse_res_content(res_content):
    # 提取 tasks_dict
    tasks_dict_str = re.search(r'tasks_dict\s*=\s*(\{.*?\})', res_content, re.S).group(1)
    tasks_dict = ast.literal_eval(tasks_dict_str)

    # ver_map
    ver_map_match = re.search(r'ver_map\s*=\s*(\[.*?\])', res_content, re.S)
    if ver_map_match:
        ver_map = ast.literal_eval(ver_map_match.group(1))
    else:
        ver_map = []

    # mode 默认 DT_HCT
    mode_match = re.search(r'mode\s*=\s*["\']?(.*?)["\']?\s*$', res_content, re.M)
    if mode_match:
        mode = mode_match.group(1)
    else:
        mode = "DT_HCT"

    return tasks_dict, ver_map, mode


def get_tenant_access_token(app_id: str, app_secret: str):
    """
    获取飞书 tenant_access_token
    """
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {
        "app_id": app_id,
        "app_secret": app_secret
    }

    try:
        resp = requests.post(url, json=payload, timeout=5)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise RuntimeError(f"请求飞书接口失败: {e}")

    data = resp.json()
    if data.get("code") != 0:
        raise RuntimeError(f"获取 tenant_access_token 失败: {data}")

    token = data["tenant_access_token"]
    expire = data["expire"]
    return token, expire


def upload_file(file_path, file_name, te_token):
    file_size = os.path.getsize(file_path)
    url = "https://open.feishu.cn/open-apis/drive/v1/files/upload_all"
    form = {'file_name': file_name,
            'parent_type': 'explorer',
            'parent_node': 'Col0fBF8slvQrFdzEfAcAWygnYs',
            'size': str(file_size),
            'file': (open(file_path, 'rb'))}
    multi_form = MultipartEncoder(form)
    headers = {
        'Authorization': f'Bearer {te_token}',
    }
    headers['Content-Type'] = multi_form.content_type
    response = requests.request("POST", url, headers=headers, data=multi_form)
    time.sleep(5)  # 等待文件上传成功
    return json.loads(response.text)["data"]["file_token"]


def transfer_sheet_to_feishu_sheet(file_token, file_name):
    request: CreateImportTaskRequest = CreateImportTaskRequest.builder() \
        .request_body(ImportTask.builder()
            .file_extension("xlsx")
            .file_token(file_token)
            .type("sheet")
            .file_name(file_name)
            .point(ImportTaskMountPoint.builder()
                .mount_type(1)
                .mount_key("Col0fBF8slvQrFdzEfAcAWygnYs")
                .build())
            .build()) \
        .build()

    response: CreateImportTaskResponse = client.drive.v1.import_task.create(request)

    if not response.success():
        lark.logger.error(
            f"client.drive.v1.import_task.create failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    time.sleep(5)  # 等待任务导入成功
    request: GetImportTaskRequest = GetImportTaskRequest.builder() \
        .ticket(response.data.ticket) \
        .build()

    response: GetImportTaskResponse = client.drive.v1.import_task.get(request)

    if not response.success():
        lark.logger.error(
            f"client.drive.v1.import_task.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    return response.data.result.token


def move_file_to_wiki(file_token, file_type):
    request: MoveDocsToWikiSpaceNodeRequest = MoveDocsToWikiSpaceNodeRequest.builder() \
        .space_id("7560534433279868930") \
        .request_body(MoveDocsToWikiSpaceNodeRequestBody.builder()
            .parent_wiki_token("GOgIwx6N8iutB5kIlXGcHuGEnnh")
            .obj_type(file_type)
            .obj_token(file_token)
            .build()) \
        .build()

    response: MoveDocsToWikiSpaceNodeResponse = client.wiki.v2.space_node.move_docs_to_wiki(request)

    if not response.success():
        lark.logger.error(
            f"client.wiki.v2.space_node.move_docs_to_wiki failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    time.sleep(5)  # 等待文件转入wiki
    task_id = response.data.task_id
    request: GetTaskRequest = GetTaskRequest.builder() \
        .task_id(task_id) \
        .task_type("move") \
        .build()

    response: GetTaskResponse = client.wiki.v2.task.get(request)

    if not response.success():
        lark.logger.error(
            f"client.wiki.v2.task.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return

    url = f"https://horizonrobotics.feishu.cn/wiki/{response.data.task.move_result[0].node.node_token}"
    return url


def make_card(sheet_url):
    # DT_HCT 模式使用 card_sv.json 模板
    with open(os.path.join(os.path.dirname(__file__), "templates", "card_sv.json"), "r", encoding="utf-8") as f:
        data = json.load(f)
    data["body"]["elements"][0]["content"] = f"[评测结果-表格]({sheet_url})"
    return json.dumps(data, ensure_ascii=False)


def process_message_async(data: P2ImMessageReceiveV1, res_content: str):
    """
    耗时任务：解析消息、跑 HCT 对比、上传文件、发消息
    """
    try:
        te_token, _ = get_tenant_access_token(APP_ID, APP_SECRET)

        logger.info(f"输入内容：{res_content}")
        tasks_dict, ver_map, mode = parse_res_content(res_content)
        if mode != "DT_HCT":
            raise ValueError(f"仅支持 mode=DT_HCT，收到: {mode}")

        file_name_suffix = time.strftime("%Y%m%d%H%M%S", time.localtime())
        msg = ""
        msg, xlsx_path, txt_path, all_ver_com, tag_lv1_ver_com, tag_all_ver_com = run_compare_hct(
            tasks_dict, ver_map, msg, file_name_suffix
        )
        xlsx_token = upload_file(xlsx_path, f"DT多版本评测结果_hct_{file_name_suffix}.xlsx", te_token)
        feish_sheet_token = transfer_sheet_to_feishu_sheet(xlsx_token, f"DT多版本评测结果_hct_{file_name_suffix}.xlsx")
        sheet_url = move_file_to_wiki(feish_sheet_token, "sheet")

        # 生成卡片内容
        content = make_card(sheet_url)

        # 回复消息（p2p 或 群聊）
        if data.event.message.chat_type == "p2p":
            request = (
                CreateMessageRequest.builder()
                .receive_id_type("chat_id")
                .request_body(
                    CreateMessageRequestBody.builder()
                    .receive_id(data.event.message.chat_id)
                    .uuid(data.header.event_id)
                    .msg_type("interactive")
                    .content(content)
                    .build()
                )
                .build()
            )
            response = client.im.v1.message.create(request)
        else:
            request = (
                ReplyMessageRequest.builder()
                .message_id(data.event.message.message_id)
                .request_body(
                    ReplyMessageRequestBody.builder()
                    .content(content)
                    .msg_type("interactive")
                    .build()
                )
                .build()
            )
            response = client.im.v1.message.reply(request)

        if not response.success():
            raise Exception(f"消息发送失败: {response.code}, {response.msg}, log_id: {response.get_log_id()}")

    except Exception as e:
        content = json.dumps(
                {
                    "text": "评测对比任务失败" + str(e)
                }
        )
        if data.event.message.chat_type == "p2p":
            request = (
                CreateMessageRequest.builder()
                .receive_id_type("chat_id")
                .request_body(
                    CreateMessageRequestBody.builder()
                    .receive_id(data.event.message.chat_id)
                    .uuid(data.header.event_id)
                    .msg_type("text")
                    .content(content)
                    .build()
                )
                .build()
            )
            response = client.im.v1.message.create(request)
        else:
            request = (
                ReplyMessageRequest.builder()
                .message_id(data.event.message.message_id)
                .request_body(
                    ReplyMessageRequestBody.builder()
                    .content(content)
                    .msg_type("text")
                    .build()
                )
                .build()
            )
            response = client.im.v1.message.reply(request)
        if not response.success():
            raise Exception(f"消息发送失败: {response.code}, {response.msg}, log_id: {response.get_log_id()}")


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
    if data.header.event_id in memory["events_id"] or time_diff(data.header.create_time):
        return
    else:
        memory["events_id"].add(data.header.event_id)
    logger.info(f"Received a new message, start processing event_id={data.header.event_id}")
    executor.submit(process_message_async, data, res_content)


# 注册事件回调
event_handler = (
    lark.EventDispatcherHandler.builder("", "")
    .register_p2_im_message_receive_v1(do_p2_im_message_receive_v1)
    .build()
)


# 创建 LarkClient 和 LarkWSClient
client = lark.Client.builder().app_id(APP_ID).app_secret(APP_SECRET).build()
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
