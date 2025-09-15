# 使用Python 3.9 官方镜像
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1

# 配置pip国内镜像源
RUN pip config set global.index-url https://mirrors.aliyun.com/pypi/simple/ && \
    pip config set global.trusted-host mirrors.aliyun.com

# 复制依赖文件并安装
COPY excel_tool/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY excel_tool/ .

# 创建必要的目录
RUN mkdir -p uploads results logs

# 暴露端口
EXPOSE 5000

# 启动命令
CMD ["python", "app.py"]