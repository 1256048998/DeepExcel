# DeepExcel Server

账号、权益、**出口路由**与遥测。模型流量**不经过**本服务。

对应路线图 [M2 — 账号体系 + 服务端地基](../docs/superpowers/specs/2026-09-12-deepexcel-roadmap-and-design.md)。

## 为什么现在建它

不是为了收费——收费是 M6。现在建有三个理由：

1. **账号是唯一越晚做越贵的东西。** 每个不带身份概念构建的客户端功能（技能库、遥测、配额）将来都要返工。
2. **现在完全没有分发控制。** 拿到安装包的人都能装，我们不知道有多少人在用。邀请码注册解决这个。
3. **遥测几乎免费搭车。** 有了服务端和用户 ID，可观测性只是多两张表；`任务一次成功率`第一次变得可测量。

## 最关键的设计：出口路由抽象

`GET /api/v1/session/endpoint` 是本期唯一不可逆的决定。

```jsonc
// 当前（HOSTED_PROXY_BASE_URL 未设置）
{ "mode": "byok",   "base_url": null, "auth_header": null, ... }

// M6（设置 HOSTED_PROXY_BASE_URL + 把权益改为 hosted）
{ "mode": "hosted", "base_url": "https://api.deepexcel.com/v1",
  "auth_header": "Bearer <短期令牌>", ... }
```

**客户端契约（两边都必须遵守）：**

- 客户端不存任何"该用哪种模式"的规则，只应用服务端返回的内容。客户端里出现 `if plan == ...` 就破坏了整个设计——而且客户端未签名、可被任意修改，客户端侧的判断本来就不是强制手段。
- `auth_header` 是完整的头部值，客户端不需要知道鉴权方案。
- 客户端在 `expires_at` 到期时重新拉取。这是 kill switch：一个 TTL 内就能收回访问权，无需发版。

**M6 切换客户端零改动**——这一条有测试守着：`tests/test_endpoint_routing.py::test_flipping_to_hosted_requires_no_client_change`。

两个刻意的行为：

- **权益配置为 hosted 但服务端没配代理地址 → 503，不静默降级为 byok。** hosted 用户本地没有配供应商 Key，悄悄给 byok 会让他们看到一堆无法诊断的鉴权错误，真正的原因（我们的部署）反而不可见。
- **配额只对 hosted 生效。** BYOK 用户直接付钱给供应商，对他们计数等于为我们没承载的流量收费。

## 本地运行

```bash
cd server
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements-dev.txt

# 开发模式用 SQLite，启动时自动建表
.venv/Scripts/python.exe -m uvicorn app.main:app --reload
```

打开 http://127.0.0.1:8000/docs。

```bash
.venv/Scripts/python.exe -m pytest tests/ -q
```

### 对着真实 Postgres 跑测试

SQLite 方便，但它会默默接受 Postgres 拒绝的东西——`audit_logs.id` 上一个多余的唯一约束
（主键本来就唯一）在 SQLite 上毫无反应，换成 Postgres 立刻报出模型与迁移不一致。所以整套
测试必须能对着真实数据库跑：

```bash
docker run -d --name deepexcel-pgtest -p 55432:5432 \
  -e POSTGRES_PASSWORD=testpw -e POSTGRES_USER=deepexcel -e POSTGRES_DB=deepexcel_test \
  postgres:16-alpine

export DEEPEXCEL_TEST_DATABASE_URL="postgresql+psycopg://deepexcel:testpw@127.0.0.1:55432/deepexcel_test"
export DATABASE_URL="$DEEPEXCEL_TEST_DATABASE_URL"
.venv/Scripts/python.exe -m pytest tests/ -q

docker rm -f deepexcel-pgtest
```

发版前至少跑一次。

## 生产部署

```bash
cp .env.example .env      # 填写 JWT_SECRET / POSTGRES_PASSWORD / BOOTSTRAP_ADMIN_*
docker compose up -d
```

`migrate` 服务会先跑 `alembic upgrade head`，`api` 等它成功后才启动。数据库不对外暴露端口，`api` 只绑定 `127.0.0.1:8000`——对外由反向代理终止 TLS。

生产环境有三条硬性校验（`app/config.py`）：未设 `JWT_SECRET`、密钥短于 32 字符、或 `DATABASE_URL` 指向 SQLite，都会拒绝启动。

### 更新源（可选）

要让已安装的客户端能自动升级，设 `UPDATE_MANIFEST_DIR` 指向一个目录，每个通道一个已签名的清单文件：

```bash
UPDATE_MANIFEST_DIR=/srv/deepexcel/updates
cp update.json /srv/deepexcel/updates/stable.json
```

`GET /api/v1/updates/latest?channel=stable` 原样返回该文件。放文件即发布，不需要重启——清单按 mtime 重读。未设该变量时该端点返回 503 并说明原因。

两个刻意的设计，改之前先想清楚：

- **这个端点不鉴权。** 登录过期或从未登录的用户同样需要拿到修复，而清单的可信度来自它的签名，不是来自谁在请求。
- **服务端不签名，也绝不能持有私钥。** 它只转发离线签好的清单。正因如此，拿下这台服务器的人能做的只是扣住更新或回放旧清单，**推不了代码**。私钥一旦上了服务端，这条性质就没了，而且没有任何测试会因此变红。

服务端只做结构校验（是不是合法 JSON、必需字段在不在），不验签——验签是客户端的事，那才是需要被说服的一方。结构校验的作用是拦住"运维复制错了文件"，否则那看起来会像一个正常工作的部署，而所有客户端已经静默停止升级。

### 知识包（可选）

模型按需读取的知识技能（`src/DeepExcel.Sidecar/knowledge/`）可以不发版就更新：离线签名后放进同一个目录。

```bash
python scripts/knowledge_pack.py build --key D:/offline/deepexcel-update.pem --out knowledge_pack.json
python scripts/knowledge_pack.py verify --pack knowledge_pack.json
cp knowledge_pack.json /srv/deepexcel/updates/knowledge_pack.json
```

`GET /api/v1/updates/knowledge` 原样返回它。客户端跟着更新检查一起拉（启动 1 分钟后、之后每 6 小时），用内置公钥验签后整包替换到 `%LOCALAPPDATA%\DeepExcel\knowledge`；侧车在同名技能的缓存版本不低于内置版本时用缓存。

- **知识包和更新清单用同一把私钥。** 知识正文进入模型上下文，等于给 agent 的指令；不签名的话，拿下服务端就多了一条提示词注入通道。
- `pack_version` 默认取当前时间戳，客户端只接受比本机更新的包，回放旧包无效。
- 包里去掉的技能会从缓存消失，安装包自带的那份仍在。WPS 没有账号 / 服务端连接，只读 Excel 端同步下来的同一个缓存目录。

技能里「用户实际遇到的报错」一节由遥测驱动：`GET /admin/api/tool-errors?days=30`（管理员）导出工具 × 报错类别 × 次数，
`python scripts/knowledge_errors.py --errors tool-errors.json` 按每个技能 frontmatter 里的 `tools:` 改写那一节并把版本号加一，
审过 diff 再 `knowledge_pack.py build` 签名发布。导出里只有工具名和固定的类别代码，不含报错原文。

**镜像必须真的构建过一次才算数。** `requirements.txt` 的版本号是从可用环境导出的，不是手写的——
手写过一次 `alembic==1.16.6`，那个版本根本不存在，直到第一次 docker build 才暴露。`httpx` 也曾
只列在 dev 依赖里（本地由 pytest 带入），生产镜像一 import 代理模块就启动失败。
`tests/test_requirements.py` 现在会检查 `app/` 里 import 的每个第三方包都在 `requirements.txt` 中。

## 数据库迁移

改了 `app/models.py` 之后**必须**生成迁移：

```bash
DATABASE_URL="sqlite:///./scratch.db" .venv/Scripts/python.exe -m alembic revision --autogenerate -m "描述这次改动"
```

`tests/test_migrations.py` 会在模型与迁移不一致时失败并指出具体的列。开发和测试用 `create_all`，所以忘记生成迁移在本地完全看不出来——只会在生产上表现为"某个列不存在"。这个测试就是那道闸门。

## 安全设计要点

| 决定 | 原因 |
| --- | --- |
| 密码用 `hashlib.scrypt`（stdlib） | 内存硬 KDF，无原生编译依赖，镜像和测试都少一层风险 |
| 管理员独立表 | 与用户表共用一张表加角色字段，意味着一个认证 bug 就能把客户变成运营 |
| JWT 分 audience（user / admin / proxy） | 代理令牌泄露不等于账号被接管；管理员令牌不能重放到用户接口 |
| Refresh token 只存 SHA-256 且一次性 | 库被脱不能重放；轮转后旧令牌立即失效 |
| 禁用用户时同时吊销 refresh token | 否则"禁用"在长达一个月内形同虚设 |
| 遥测白名单在**服务端**执行 | 客户端未签名可改，它的过滤是便利而非保证 |
| 所有管理操作写审计日志 | "是谁禁用了这个账号"需要有答案 |

## 遥测隐私

绝不采集：单元格内容、工作簿名与路径、用户输入原文、模型回复、API Key、原始错误消息。

白名单在 `app/routers/telemetry.py`，并通过 `GET /api/v1/telemetry/schema` **公开**——让用户可以自己核对，而不是要求他们相信我们。白名单之外的字段在入库前被丢弃，且丢弃行为会在响应里报告（不静默）。

`tests/test_telemetry.py::test_fields_outside_the_allowlist_never_reach_storage` 用真实的工作簿路径、单元格值和 API Key 作为输入，断言它们不会出现在存储中。

## 尚未实现

- **支付渠道接入**。微信支付 / 支付宝都需要企业商户号，申请周期按月计。订单、权益激活、
  续费延长、回调幂等都已实现并有测试，缺的只是最后那个"谁把订单标记为已支付"——目前由
  运营在后台手工标记，每次写审计日志。接上回调后下游一行不用改。
- **上游供应商配置**。代理本身已完成，但 `UPSTREAM_*` 环境变量需要填真实的供应商密钥。
  未配置时 `/v1/messages` 返回 502 `no_upstream`，不会静默失败。

## 已验证到什么程度

- 100 项测试，SQLite 与**真实 Postgres** 两种后端都跑通
- `docker compose` 已实际构建并启动；生产配置下逐条验证过 HTTP 行为：
  `/docs` 关闭、未授权管理 API 401、无邀请码注册 403、代理未认证 401
- 端到端真实链路：管理员登录 → 生成邀请码 → 凭码注册 → 拿到出口配置 → 邀请码用完后被拒
- 迁移链在两种数据库上都可正向应用、可回滚，且与模型无漂移

`consume-task` 目前由客户端上报，可伪造。这是刻意的：在没有钱依赖它之前，先把计数、
周期重置和额度耗尽的响应跑通。代理上线后它成为权威计数点，因为那是客户端绕不过去的地方。
