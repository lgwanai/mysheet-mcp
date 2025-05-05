#!/bin/bash

# WolinkServer应用管理脚本
# 支持start|stop|restart|status命令

APP_NAME="mysheet-mcp"
APP_MAIN_CLASS="link.wo.mysheetmcp.MysheetMcpApplication"
JAR_FILE=$(ls *.jar | head -n 1)
PID_FILE="mysheet-mcp.pid"
LOG_FILE="mysheet-mcp.log"

# 检查Java环境
check_java() {
    if ! command -v java &> /dev/null; then
        echo "错误: Java未安装，请先安装Java"
        exit 1
    fi
}

# 检查jar文件
check_jar() {
    if [ -z "$JAR_FILE" ]; then
        echo "错误: 未找到可执行的jar文件，请先构建项目(mvn clean package)"
        echo "提示: 如果jar文件不在target目录，请手动设置JAR_FILE环境变量"
        exit 1
    fi
}

# 获取应用PID
get_pid() {
    if [ -f "$PID_FILE" ]; then
        PID=$(cat "$PID_FILE")
        if ps -p $PID > /dev/null; then
            echo $PID
            return
        fi
    fi
    PID=$(ps -ef | grep "$APP_MAIN_CLASS" | grep -v grep | awk '{print $2}')
    echo $PID
}

# 启动应用
start() {
    check_java
    check_jar

    PID=$(get_pid)
    if [ -n "$PID" ]; then
        echo "$APP_NAME 已经在运行(PID: $PID)"
        exit 0
    fi

    echo "正在启动 $APP_NAME..."
    nohup java -jar "$JAR_FILE" --spring.profiles.active=prod > "$LOG_FILE" 2>&1 &
    echo $! > "$PID_FILE"
    echo "$APP_NAME 启动成功(PID: $!)"
}

# 停止应用
stop() {
    PID=$(get_pid)
    if [ -z "$PID" ]; then
        echo "$APP_NAME 未在运行"
        return
    fi

    echo "正在停止 $APP_NAME(PID: $PID)..."
    kill $PID
    rm -f "$PID_FILE"
    echo "$APP_NAME 已停止"
}

# 重启应用
restart() {
    stop
    sleep 2
    start
}

# 查看状态
status() {
    PID=$(get_pid)
    if [ -z "$PID" ]; then
        echo "$APP_NAME 未在运行"
    else
        echo "$APP_NAME 正在运行(PID: $PID)"
    fi
}

# 主逻辑
case "$1" in
    start)
        start
        ;;
    stop)
        stop
        ;;
    restart)
        restart
        ;;
    status)
        status
        ;;
    *)
        echo "用法: $0 {start|stop|restart|status}"
        exit 1
        ;;
esac

exit 0
