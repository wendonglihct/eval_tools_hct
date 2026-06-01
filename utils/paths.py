"""项目路径统一管理。

所有持久化产物 / 配置 / 模板的目录都在这里集中声明，业务代码只 import 常量或
辅助函数，避免在多个文件里散落 os.path.join("...", "configs", ...) 这种硬编码。

布局:
    <PROJECT_ROOT>/
        configs/          <- CONFIG_DIR     (所有 JSON：modes.json / hct_report_config.json / card_sv.json)
        outputs/
            tasks/        <- TASKS_DIR      (pncops 下载缓存)
            reports/      <- REPORTS_DIR    (xlsx / txt 报告)
            runtime/      <- RUNTIME_DIR    (pid / stdout 日志)
            log/          <- LOG_DIR        (bot 业务日志，原 robot_memory/log)
"""

import os

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# 配置（路由表 / 报告特性 / 卡片模板等所有 JSON 都在此目录）
CONFIG_DIR = os.path.join(PROJECT_ROOT, "configs")

# 运行时产物
OUTPUTS_DIR = os.path.join(PROJECT_ROOT, "outputs")
TASKS_DIR = os.path.join(OUTPUTS_DIR, "tasks")
REPORTS_DIR = os.path.join(OUTPUTS_DIR, "reports")
RUNTIME_DIR = os.path.join(OUTPUTS_DIR, "runtime")
LOG_DIR = os.path.join(OUTPUTS_DIR, "log")

# 关键配置文件名（路由表本身的位置固定，其他配置由路由表指向）
MODE_ROUTING_FILE = "modes.json"


def config_path(name: str) -> str:
    """按名称定位 CONFIG_DIR 下的 JSON 文件（含特性配置、卡片模板等）。"""
    return os.path.join(CONFIG_DIR, name)


def ensure_output_dirs() -> None:
    """创建所有运行时产物目录（幂等）。"""
    for d in (TASKS_DIR, REPORTS_DIR, RUNTIME_DIR, LOG_DIR):
        os.makedirs(d, exist_ok=True)
