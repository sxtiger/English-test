#!/bin/sh
# 在 NAS 上运行:拉取最新代码并重启容器。
# 代码通过 .:/app 卷挂载,重启即生效,无需重建镜像。
# 仅当 requirements.txt / dockerfile 变化时需改用: docker-compose up -d --build
set -e
cd "$(dirname "$0")/.."

echo ">>> 拉取最新代码 ..."
if ! git pull --ff-only; then
    echo ""
    echo ">>> git pull 失败:NAS 本地的受跟踪文件被修改过"
    echo ">>> (常见原因:通过网页上传功能更新了短语库)"
    echo ">>> 若要保留 NAS 上的改动:  先在本机 git add/commit/push"
    echo ">>> 若要放弃 NAS 上的改动:  git checkout -- . && sh scripts/nas_update.sh"
    exit 1
fi

echo ">>> 重启容器 ..."
docker restart english-phrase-test

echo ">>> 完成。数秒后可查看状态:"
echo "    docker ps | grep english-phrase-test"
