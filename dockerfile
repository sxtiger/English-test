# 使用Python 3.9官方镜像
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# openssl 用于首次启动时生成自签名证书
RUN apt-get update && apt-get install -y --no-install-recommends \
    openssl \
    && rm -rf /var/lib/apt/lists/*

# 复制requirements文件
COPY requirements.txt .

# 预装Python依赖（构建时完成，容器启动不再依赖网络，秒级启动）
RUN pip install --no-cache-dir -r requirements.txt

# 创建必要的目录
RUN mkdir -p /app/flask_sessions /app/uploads /app/templates /app/static /app/ssl

# 启动脚本放在 /usr/local/bin，避免被 .:/app 卷挂载覆盖
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh
RUN chmod +x /usr/local/bin/docker-entrypoint.sh

# 暴露端口：51000=HTTP，51001=HTTPS
EXPOSE 51000 51001

# 健康检查（slim镜像不含curl，改用python探测）
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:51000/health', timeout=5)" || exit 1

ENTRYPOINT ["docker-entrypoint.sh"]
CMD ["python", "exam_webui_https.py"]
