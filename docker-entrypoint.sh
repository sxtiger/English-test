#!/bin/sh
set -e

cd /app

# 确保运行所需目录存在
mkdir -p flask_sessions uploads data ssl

# 首次启动自动生成自签名证书（有效期10年）
# 若需更换证书（如NAS的IP变更），删除 ./ssl 目录后重启容器即可重新生成
if [ ! -f ssl/cert.pem ] || [ ! -f ssl/privkey.pem ]; then
    echo ">>> 首次启动，生成自签名SSL证书 (服务器地址: ${SERVER_IP:-127.0.0.1}) ..."
    openssl req -x509 -newkey rsa:2048 \
        -keyout ssl/privkey.pem \
        -out ssl/cert.pem \
        -days 3650 -nodes \
        -subj "/CN=${SERVER_IP:-127.0.0.1}" \
        -addext "subjectAltName=DNS:localhost,IP:${SERVER_IP:-127.0.0.1},IP:127.0.0.1"
    echo ">>> 证书生成完成: /app/ssl/cert.pem"
fi

exec "$@"
