# eval_tools_hct
./start-robot.sh start	检查 PID 文件是否已运行 → 不在则 nohup python -u main.py 后台启动，PID 写入文件，1 秒后探活
./start-robot.sh stop	读取 PID，SIGTERM 优雅停 5 秒；超时 SIGKILL；清理 PID
./start-robot.sh restart	stop + start
./start-robot.sh status	显示 running/not running + ps 资源信息 + 日志路径
./start-robot.sh logs [bot|out]	tail -f 业务日志或 stdout；默认 bot
./start-robot.sh help	用法说明