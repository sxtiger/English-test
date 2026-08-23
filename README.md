# 英语短语测试系统(English Phrase Test)

基于 Flask 的英语短语四选一测试系统,部署于内网 Synology NAS 的 Docker 中,同时提供 HTTP 与 HTTPS 访问。支持正常测试与错题重练两种模式,短语库以 Markdown 文件维护,可通过网页直接上传更新。

## 功能特性

- **双协议访问**:同一进程同时提供 HTTP(51000)与 HTTPS(51001)真实服务
- **两种测试模式**:正常模式(全库随机抽题)、错题模式(仅从错题库抽题)
- **双向出题**:英文选中文 / 中文选英文,随机切换
- **错题自动积累**:正常模式答错的短语自动写入 `Error_Phrases.xlsx`
- **短语库网页上传**:上传新的 `.md` 短语库,旧库自动备份为 `English Phrase_backup_*.md`
- **容器化部署**:依赖在构建镜像时预装,容器秒级启动、离线可运行,健康检查 `/health`
- **自签名证书自动生成**:首次启动自动生成 10 年期证书(SAN 含 NAS 的 IP)

## 技术栈

Python 3.9 · Flask 2.3 · Flask-Session · pandas · openpyxl · Docker

## 短语库格式

`English Phrase.md` 是唯一的短语库源文件,格式如下:

```markdown
# 书本名称
## Day 1
- hello: 你好
- world: 世界

## Day 2
- good morning: 早上好
```

一级标题为书本名称,二级标题为学习日,列表项 `- 英文: 中文` 为短语条目。系统每次加载首页时会自动将其转换为 `English Phrase.xlsx`(运行时生成,不纳入 git)。

## 目录结构

```
.
├── exam_webui_https.py      # 主程序入口(HTTP + HTTPS 双协议)
├── md2excel.py              # Markdown → Excel 转换工具
├── English Phrase.md        # 短语库源文件(纳入版本管理)
├── templates/               # 页面模板(index / test / result / error)
├── requirements.txt         # Python 依赖
├── dockerfile               # 镜像构建文件(依赖预装)
├── docker-compose.yml       # 部署编排(当前使用)
├── docker-entrypoint.sh     # 容器启动脚本(建目录、生成证书)
├── scripts/nas_update.sh    # NAS 端一键更新脚本(git pull + 重启容器)
├── .env.example             # 环境变量模板(复制为 .env 填写内网地址;.env 不发布)
│
│  ── 以下为本地/运行时文件,已在 .gitignore 中排除,不发布 ──
├── 对话记录_*.md、改动与部署说明_*.md # 本地工作记录,仅本地留存
├── English Phrase.xlsx      # 由 .md 自动转换的题库
├── Error_Phrases.xlsx       # 运行时积累的错题库
├── English Phrase_backup_*.md # 上传新短语库时的自动备份
├── flask_sessions/          # Flask 会话文件
├── ssl/                     # 自签名证书与私钥(严禁外传)
└── uploads/、data/          # 上传临时目录 / 预留目录
```

> `exam_webui.py`(旧版单 HTTP 入口)与 `docker-compose-https.yml`(旧版启动时装依赖方案)**已废弃**,仅作存档保留,新部署请勿使用。

## 访问地址

| 协议  | 地址 |
|-------|------|
| HTTP  | `http://<NAS_IP>:51000` |
| HTTPS | `https://<NAS_IP>:51001` |

> 下文中的 `<NAS_IP>` 均指你自己 NAS 的局域网 IP。

HTTPS 使用自签名证书,浏览器首次访问提示"不安全",选择"高级 → 继续访问"即可。

## 本地运行

```bash
pip install -r requirements.txt

# 首次运行需先生成自签名证书(只需一次)
mkdir -p ssl
openssl req -x509 -newkey rsa:2048 \
    -keyout ssl/privkey.pem -out ssl/cert.pem \
    -days 3650 -nodes \
    -subj "/CN=127.0.0.1" \
    -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

python exam_webui_https.py
```

## Docker 部署(Synology NAS)

部署目录约定为 `/volume1/docker/english-test`,容器名 `english-phrase-test`。

### 首次部署

**① 配置本机地址**:在项目目录创建 `.env`(不发布到 GitHub),填入 NAS 的局域网 IP:

```bash
cp .env.example .env
# 编辑 .env,填写如:SERVER_IP=192.168.1.100
```

**② 构建并启动**

**方式一:Container Manager 图形界面(推荐,DSM 7.2)**

1. 打开 **Container Manager → 项目 → 新增**
2. 项目名称随意(如 `english-test`),路径选择 `docker/english-test` 共享文件夹
3. 系统自动读取其中的 `docker-compose.yml`,构建镜像并启动容器

**方式二:SSH 命令行**

```bash
cd /volume1/docker/english-test
docker build -f dockerfile -t english-phrase-test:latest .
docker-compose up -d
docker ps            # 状态应变为 healthy
docker logs -f english-phrase-test
```

### 日常更新(实时部署)

代码与模板通过 `.: /app` 卷挂载进容器,**修改同步到 NAS 后重启容器即生效(约 3 秒),无需重建镜像**。完整流程见下文[开发与部署流程](#开发与部署流程git)。

仅当 `requirements.txt` 或 `dockerfile` 变化时才需要重建镜像:

```bash
docker-compose up -d --build
```

## 开发与部署流程(git)

本地编辑 → 提交推送 GitHub → NAS 拉取并重启容器:

```bash
# ① 本地(修改代码后)
git add -A
git commit -m "描述本次改动"
git push

# ② NAS(SSH 登录后)
cd /volume1/docker/english-test
sh scripts/nas_update.sh     # 等价于 git pull --ff-only && docker restart english-phrase-test
```

NAS 端一次性初始化配置见 [NAS 端初始化](#nas-端一次性初始化)。

**注意**:启用 git 流程后,短语库请尽量通过"本地编辑 → 部署"更新。若仍使用网页上传功能,上传会直接改写 NAS 上的 `English Phrase.md`,导致下次 `git pull` 被拒绝;此时需在 NAS 上提交推送该改动,或用 `git checkout -- .` 放弃后再拉取。

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `SERVER_IP` | `127.0.0.1` | NAS 地址,写入自签名证书的 SAN;请填入项目目录的 `.env` 文件(参见 `.env.example`) |
| `HTTP_PORT` | `51000` | HTTP 端口 |
| `HTTPS_PORT` | `51001` | HTTPS 端口 |
| `SECRET_KEY` | 内置默认值 | Flask 会话密钥,可在 `.env` 或 `docker-compose.yml` 中覆盖 |

## NAS 端初始化

前置条件:

1. DSM **控制面板 → 终端机和 SNMP → 终端** 中启用 SSH
2. **套件中心** 安装社区套件 **Git Server**(提供 git 命令)
3. 已有 GitHub 私有仓库及访问凭据(个人访问令牌 PAT 或 SSH 密钥)

在 NAS 上执行(将现有部署目录初始化为 git 工作副本):

```bash
cd /volume1/docker/english-test
git init -b main
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git fetch origin
git reset --hard origin/main   # ⚠️ 会以仓库版本覆盖同名文件;执行前确认本地短语库已是最新
git branch -u origin/main

# 配置本机局域网地址(证书生成与启动日志使用;.env 不会发布)
cp .env.example .env
vi .env                        # 填写 SERVER_IP=<NAS_IP>
```

运行时文件(`ssl/`、`flask_sessions/`、各 `.xlsx`、备份文件)均不在 git 跟踪范围内,上述操作不会影响它们。

之后每次更新只需:`sh scripts/nas_update.sh`

> 若 git 提示 "detected dubious ownership",执行:
> `git config --global --add safe.directory /volume1/docker/english-test`

## 常见问题

1. **局域网无法访问**:若启用了 DSM 防火墙(控制面板 → 安全性 → 防火墙),需放行 TCP `51000` 与 `51001`。
2. **NAS 更换 IP**:修改 `.env` 中的 `SERVER_IP` 为新 IP → 删除项目目录下的 `ssl/` → 重启容器,证书按新 IP 重新生成。
3. **数据持久性**:短语库、错题库、会话均保存在项目目录(卷挂载),重建容器不丢数据;但错题库不在 git 中,换机全新部署时从零开始。
4. **确认部署成功**:容器状态为 `healthy`,日志出现:
   ```
   === 启动测试系统（HTTP + HTTPS 双协议）===
   HTTP  访问地址: http://<NAS_IP>:51000
   HTTPS 访问地址: https://<NAS_IP>:51001
   ```

## 更新日志

- **2026-08-23** 修复容器启动后局域网无法访问的问题;重写 Dockerfile(依赖预装、秒级启动);HTTP/HTTPS 双协议同进程服务;新增健康检查与自签名证书自动生成(详见本地文档《改动与部署说明_20260823.md》,不随仓库发布);初始化 git 仓库与 `.gitignore`,建立 GitHub → NAS 的 git 部署流程;发布内容脱敏:`SERVER_IP` 改由本地 `.env` 配置,文档不再包含具体内网地址
- **2025-09-16** 优化代码与 HTML 模板;支持 http/https;引入 docker compose
