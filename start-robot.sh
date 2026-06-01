#!/bin/bash
# 启动 / 停止 / 重启 飞书机器人（bot/robot.py 经 main.py 入口）
# 用法:
#   ./start-robot.sh start|stop|restart|status|logs|help

set -e

CYAN="\033[0;36m"; GREEN="\033[0;32m"; YELLOW="\033[1;33m"
GRAY="\033[90m"; RED="\033[0;31m"; BOLD="\033[1m"; NC="\033[0m"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENTRY="$SCRIPT_DIR/main.py"

RUNTIME_DIR="$SCRIPT_DIR/outputs/runtime"
PID_FILE="$RUNTIME_DIR/robot.pid"
STDOUT_LOG="$RUNTIME_DIR/robot.out"
BOT_LOG="$SCRIPT_DIR/outputs/log/chatbot.log"

mkdir -p "$RUNTIME_DIR"

# 取一个可用的 python 解释器（优先工程 venv）
resolve_python() {
    if [ -x "$SCRIPT_DIR/.venv/bin/python" ]; then
        echo "$SCRIPT_DIR/.venv/bin/python"
    elif command -v python3 >/dev/null 2>&1; then
        echo "python3"
    else
        echo "python"
    fi
}

# 校验 PID 是否真在跑 main.py（避免 PID 复用）
is_running() {
    [ -f "$PID_FILE" ] || return 1
    local pid; pid=$(cat "$PID_FILE" 2>/dev/null || true)
    [ -z "$pid" ] && return 1
    kill -0 "$pid" 2>/dev/null || return 1
    # 进一步校验 cmdline 中包含 main.py
    if [ -r "/proc/$pid/cmdline" ]; then
        tr '\0' ' ' < "/proc/$pid/cmdline" | grep -q "main.py" || return 1
    fi
    return 0
}

cmd_start() {
    if is_running; then
        local pid; pid=$(cat "$PID_FILE")
        echo -e "${YELLOW}● robot 已在运行${NC}  pid=${BOLD}$pid${NC}"
        return 0
    fi
    rm -f "$PID_FILE"

    local py; py=$(resolve_python)
    echo -e "${CYAN}启动 robot${NC}  python=${py}  entry=${ENTRY}"
    cd "$SCRIPT_DIR"
    nohup "$py" -u "$ENTRY" >>"$STDOUT_LOG" 2>&1 &
    local pid=$!
    echo "$pid" > "$PID_FILE"

    # 给一点时间检查是否秒退
    sleep 1
    if kill -0 "$pid" 2>/dev/null; then
        echo -e "${GREEN}✓ robot 已启动${NC}  pid=${BOLD}$pid${NC}"
        echo -e "  ${GRAY}stdout/stderr → $STDOUT_LOG${NC}"
        echo -e "  ${GRAY}bot 业务日志   → $BOT_LOG${NC}"
    else
        echo -e "${RED}✗ robot 启动后立即退出，请检查 $STDOUT_LOG${NC}"
        rm -f "$PID_FILE"
        exit 1
    fi
}

cmd_stop() {
    if ! is_running; then
        echo -e "${GRAY}● robot 未运行${NC}"
        rm -f "$PID_FILE"
        return 0
    fi
    local pid; pid=$(cat "$PID_FILE")
    echo -e "${CYAN}停止 robot${NC}  pid=${BOLD}$pid${NC}"
    kill "$pid" 2>/dev/null || true
    # 等待最多 5s
    for _ in 1 2 3 4 5; do
        kill -0 "$pid" 2>/dev/null || break
        sleep 1
    done
    if kill -0 "$pid" 2>/dev/null; then
        echo -e "${YELLOW}⚠ 优雅停止失败，强制 kill -9${NC}"
        kill -9 "$pid" 2>/dev/null || true
    fi
    rm -f "$PID_FILE"
    echo -e "${GREEN}✓ 已停止${NC}"
}

cmd_restart() {
    cmd_stop
    cmd_start
}

cmd_status() {
    if is_running; then
        local pid; pid=$(cat "$PID_FILE")
        echo -e "${GREEN}● running${NC}  pid=${BOLD}$pid${NC}"
        ps -o pid,etime,pcpu,pmem,cmd -p "$pid" 2>/dev/null | sed 's/^/  /'
    else
        echo -e "${GRAY}○ not running${NC}"
    fi
    echo -e "  ${GRAY}stdout/stderr: $STDOUT_LOG${NC}"
    echo -e "  ${GRAY}bot log      : $BOT_LOG${NC}"
}

cmd_logs() {
    local target="${1:-bot}"
    case "$target" in
        bot|"")  exec tail -f "$BOT_LOG" ;;
        out|stdout) exec tail -f "$STDOUT_LOG" ;;
        *) echo -e "${RED}未知日志: $target（可选 bot|out）${NC}"; exit 1 ;;
    esac
}

cmd_help() {
    cat <<EOF
用法: $0 <命令> [参数]

命令:
  start            启动 robot（后台 nohup）
  stop             停止 robot
  restart          重启 robot
  status           查看运行状态 / PID / 资源
  logs [bot|out]   实时查看日志（默认 bot 业务日志；out 为标准输出）
  help             显示帮助

文件:
  PID    : $PID_FILE
  stdout : $STDOUT_LOG
  bot log: $BOT_LOG
EOF
}

CMD="${1:-help}"
shift || true
case "$CMD" in
    start)   cmd_start ;;
    stop)    cmd_stop ;;
    restart) cmd_restart ;;
    status)  cmd_status ;;
    logs)    cmd_logs "$@" ;;
    help|-h|--help) cmd_help ;;
    *) echo -e "${RED}未知命令: $CMD${NC}"; cmd_help; exit 1 ;;
esac
