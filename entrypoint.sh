#!/bin/bash

# 1. 准备环境
if [ -f "bin/activate" ]; then
    . bin/activate
fi

# 2. 安装依赖
echo "正在安装依赖..."
pip install -r "docx服务/requirements.txt"

# 3. 启动定时清理服务 (后台运行)
echo "正在启动定时清理服务 (12h间隔)..."
python3 cleanup_service.py > cleanup.log 2>&1 &
CLEANUP_PID=$!

# 4. 启动主 Web 服务
echo "正在启动 Docx 主服务..."
cd "docx服务"

# 捕获退出信号以清理子进程
trap 'kill $CLEANUP_PID; exit' SIGINT SIGTERM

python3 main.py

