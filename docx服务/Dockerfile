# 使用轻量级的 Python 镜像
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 安装系统依赖（如 libxml2 等，如果需要）
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件并安装
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制整个项目（包含解析器、生成器和服务逻辑）
COPY . .

# 暴露端口（FastAPI 默认端口）
EXPOSE 8080

# 启动命令
CMD ["python", "main.py"]
