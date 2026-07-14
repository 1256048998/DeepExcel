# DeepExcel SaaS 服务端 MVP

**日期：** 2026-07-14
**状态：** 设计已完成，待实现
**关联：** 本 spec 只覆盖服务端。客户端改造（登录 UI + sidecar 启动流程）将在后续 spec 处理。

## 背景与动机

DeepExcel 当前是纯本地客户端——用户自己填 Claude API Key，无账号、无订阅、无服务端。
要 SaaS 化，需要引入用户后台管理系统，支持订阅付费。

**核心架构决策（用户确认）：**
- Agent loop（Claude Agent SDK）留在客户端本地——工具调用零延迟，使用用户电脑本地能力
- 服务端只做"门卫 + 日志"，不参与 agent loop 编排
- 客户端启动时用 JWT 拉取 Key，存内存不落盘

**分阶段安全策略：**

| 阶段 | Key 下发方式 | 适用场景 | 风险 |
|------|------------|---------|------|
| **Phase 1（MVP）** | 服务端直接下发真实 Key | 早期验证、目标用户是 Excel 普通用户 | Key 可被技术手段提取（内存/抓包/改代码） |
| **Phase 2（成熟期）** | 短期 token + 服务端代理 | 用户含技术人员、企业版 | 增加带宽成本、SPOF、延迟 |

**Phase 1 风险评估：**
- 目标用户是 Excel 普通用户，不会用 Process Hacker 抓内存、不会用 Fiddler 抓包
- 风险可接受，优先验证商业模式
- 设计预留演进路径：客户端 API 不变，后续切代理只需服务端加代理端点 + 客户端改 base_url

**Phase 2 演进路径（不在本 spec 范围）：**
- 客户端登录后拿短期 token（15 分钟），不是真实 Key
- sidecar 的 Claude Agent SDK 配置 base_url 指向服务端代理端点，api_key 用短期 token
- 服务端代理层校验 token + 订阅 → 用真实 Key 透传到 Anthropic API
- 真实 Key 永远不出服务端
- 代理必须透明转发（不改 body、不改 header），否则破坏 Anthropic prompt caching

**MVP 范围（用户确认）：**
- 用户注册/登录（邮箱密码）
- 订阅状态校验（30 天试用）
- API Key 下发
- 用户日志上报（方便修 bug）
- 不做用量统计、不做计费、不做管理后台 UI

## 架构

```
┌─────────────────────────────────────────────────────────┐
│ Excel 插件（客户端，改动小）                             │
│                                                          │
│  ┌──────────┐  ┌──────────────────┐  ┌───────────────┐  │
│  │ 登录 UI   │  │ sidecar.py       │  │ 现有功能不变  │  │
│  │ 邮箱密码  │  │ Agent SDK 留本地 │  │ 工具调用本地  │  │
│  │ JWT 存储  │  │ 用服务端下发的Key│  │ 零延迟        │  │
│  └─────┬────┘  └────────┬─────────┘  └───────────────┘  │
│        │                │                                  │
│        │ ①登录          │ ③启动时用JWT拉取Key              │
│        ▼                ▼                                  │
└────────┼────────────────┼─────────────────────────────────┘
         │                │
         ▼                ▼
┌─────────────────────────────────────────────────────────┐
│ DeepExcel Server (FastAPI + Docker Compose)             │
│                                                          │
│  ① POST /api/auth/register  — 注册                     │
│  ② POST /api/auth/login     — 登录，返回 JWT           │
│  ③ GET  /api/key            — JWT+订阅校验→下发Key     │
│  ④ POST /api/logs           — 接收客户端日志（修bug）  │
│  ⑤ GET  /api/subscription   — 查订阅状态               │
│                                                          │
│  ┌─────────┐ ┌──────────┐ ┌──────────┐ ┌────────────┐ │
│  │Auth     │ │Key下发   │ │订阅校验  │ │Logger      │ │
│  │bcrypt   │ │内存缓存  │ │到期/冻结 │ │用户操作日志│ │
│  │JWT签发  │ │Key轮换   │ │          │ │            │ │
│  └─────────┘ └──────────┘ └──────────┘ └────────────┘ │
│                                                          │
│  PostgreSQL              Redis                          │
│  用户/订阅/日志           JWT黑名单/Key缓存              │
└─────────────────────────────────────────────────────────┘
```

**核心原则：Agent loop 留本地，服务端只做"门卫 + 日志"。**

数据流：
1. 用户在 Excel 插件登录 → 服务端校验 → 返回 JWT
2. 客户端启动 sidecar 前，用 JWT 调 `/api/key` → 服务端校验订阅有效 → 返回 Claude API Key（存内存，不落盘）
3. sidecar 用拿到的 Key 启动 Claude Agent SDK，后续 AI 调用全在本地
4. 客户端异步把操作日志上报到 `/api/logs`（不阻塞 AI 调用）

## 目录结构

```
server/                          ← 新增，与 src/ 平级
├── docker-compose.yml           ← FastAPI + PostgreSQL + Redis
├── Dockerfile
├── requirements.txt
├── .env.example                 ← CLAUDE_API_KEY, JWT_SECRET, DB_URL 等
├── app/
│   ├── main.py                  ← FastAPI 入口
│   ├── config.py                ← 环境变量配置
│   ├── auth/
│   │   ├── router.py            ← /api/auth/*
│   │   ├── service.py           ← 注册/登录逻辑
│   │   └── jwt.py               ← JWT 签发/校验
│   ├── key/
│   │   └── router.py            ← /api/key
│   ├── subscription/
│   │   ├── router.py            ← /api/subscription
│   │   └── service.py           ← 订阅状态逻辑
│   ├── logs/
│   │   ├── router.py            ← /api/logs
│   │   └── models.py            ← LogRecord 数据模型
│   └── db/
│       ├── models.py            ← User, Subscription, LogRecord
│       └── database.py          ← SQLAlchemy 连接
├── alembic/                     ← 数据库迁移
│   ├── alembic.ini
│   └── versions/
├── tests/
│   ├── test_auth.py
│   ├── test_key.py
│   ├── test_subscription.py
│   └── test_logs.py
└── README.md
```

## API 设计

### 认证

#### POST /api/auth/register
注册新用户。

**请求：**
```json
{
  "email": "user@example.com",
  "password": "securepassword"
}
```

**响应 201：**
```json
{
  "user_id": "uuid",
  "email": "user@example.com",
  "token": "jwt_token",
  "expires_at": "2026-07-21T00:00:00Z"
}
```

**错误：**
- 400: 邮箱格式无效
- 409: 邮箱已注册

**逻辑：**
- bcrypt 哈希密码（cost=12）
- 创建 User 记录
- 创建 Subscription 记录（tier=trial, status=active, expires_at=now+30d）
- 签发 JWT（含 user_id, tier, exp）
- 注册成功自动登录（返回 token）

#### POST /api/auth/login
登录。

**请求：**
```json
{
  "email": "user@example.com",
  "password": "securepassword"
}
```

**响应 200：**
```json
{
  "user_id": "uuid",
  "email": "user@example.com",
  "token": "jwt_token",
  "expires_at": "2026-07-21T00:00:00Z"
}
```

**错误：**
- 401: 邮箱或密码错误

### Key 下发

#### GET /api/key
获取 Claude API Key。需 JWT 认证。

**请求头：** `Authorization: Bearer <jwt_token>`

**响应 200：**
```json
{
  "api_key": "sk-ant-...",
  "provider": "anthropic",
  "base_url": "https://api.anthropic.com"
}
```

**错误：**
- 401: JWT 无效/过期
- 403: 订阅过期/冻结

**逻辑：**
- 校验 JWT
- 查订阅状态：status=active 且 expires_at > now
- 有效：返回服务端配置的 Claude API Key（从环境变量读）
- 无效：返回 403

### 订阅

#### GET /api/subscription
查订阅状态。需 JWT 认证。

**响应 200：**
```json
{
  "tier": "trial",
  "status": "active",
  "expires_at": "2026-08-14T00:00:00Z",
  "days_remaining": 25
}
```

**错误：**
- 401: JWT 无效/过期

### 日志

#### POST /api/logs
上报客户端日志。需 JWT 认证。支持批量上报。

**请求：**
```json
{
  "logs": [
    {
      "timestamp": "2026-07-14T10:30:00Z",
      "level": "info",
      "event": "sidecar_start",
      "message": "sidecar started, model=claude-sonnet-4-5",
      "context": {"model": "claude-sonnet-4-5", "session_id": "abc123"}
    },
    {
      "timestamp": "2026-07-14T10:31:00Z",
      "level": "warning",
      "event": "tool_call",
      "message": "read_range returned 500 rows",
      "context": {"tool": "read_range", "rows": 500}
    }
  ]
}
```

**响应 200：**
```json
{
  "accepted": 2
}
```

**逻辑：**
- 批量插入 LogRecord（user_id 从 JWT 提取）
- 脱敏：服务端不校验内容，但客户端上报前应去掉 API Key、密码、Excel 敏感数据

## 数据库模型

```python
# server/app/db/models.py

class User(Base):
    __tablename__ = "users"
    id = Column(UUID, primary_key=True, default=uuid4)
    email = Column(String(255), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)  # bcrypt
    created_at = Column(DateTime, default=datetime.utcnow)
    subscription = relationship("Subscription", back_populates="user", uselist=False)
    logs = relationship("LogRecord", back_populates="user")

class Subscription(Base):
    __tablename__ = "subscriptions"
    id = Column(UUID, primary_key=True, default=uuid4)
    user_id = Column(UUID, ForeignKey("users.id"), unique=True, nullable=False)
    tier = Column(String(32), default="trial")  # trial | basic | pro
    status = Column(String(32), default="active")  # active | expired | frozen
    expires_at = Column(DateTime, nullable=False)  # now + 30d
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    user = relationship("User", back_populates="subscription")

class LogRecord(Base):
    __tablename__ = "logs"
    id = Column(UUID, primary_key=True, default=uuid4)
    user_id = Column(UUID, ForeignKey("users.id"), nullable=False, index=True)
    timestamp = Column(DateTime, nullable=False, index=True)
    level = Column(String(16), default="info")  # info | warning | error
    event = Column(String(64))  # sidecar_start | tool_call | error 等
    message = Column(Text)
    context = Column(JSON)  # 附加数据
    created_at = Column(DateTime, default=datetime.utcnow)  # 服务端接收时间
    user = relationship("User", back_populates="logs")
```

**索引：**
- users.email — 唯一索引，登录查询
- subscriptions.user_id — 唯一索引，订阅查询
- logs(user_id, timestamp) — 复合索引，按用户查日志

## 安全考量

- **API Key 不落盘**：客户端拿到 Key 只存内存，sidecar 进程退出即销毁
- **JWT 加密存储**：客户端用 DPAPI（现有 SecurityManager）加密存 JWT
- **服务端 Key 轮换**：服务端可随时换 Claude API Key（改环境变量），客户端下次启动自动拉新 Key
- **HTTPS**：生产环境必须 HTTPS，JWT 和 Key 传输加密
- **日志脱敏**：客户端上报前去掉 API Key、密码、Excel 敏感数据
- **密码哈希**：bcrypt cost=12
- **JWT 过期**：7 天有效期，客户端可调 `/api/auth/refresh` 续期（MVP 可选，先不做）
- **Rate Limiting**：登录接口限流（同 IP 每分钟 10 次），防止暴力破解

## 部署

### docker-compose.yml

```yaml
version: "3.8"
services:
  api:
    build: .
    ports:
      - "8000:8000"
    env_file: .env
    depends_on:
      - db
      - redis
    restart: unless-stopped

  db:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: deepexcel
      POSTGRES_USER: deepexcel
      POSTGRES_PASSWORD: ${DB_PASSWORD}
    volumes:
      - pgdata:/var/lib/postgresql/data
    restart: unless-stopped

  redis:
    image: redis:7-alpine
    restart: unless-stopped

volumes:
  pgdata:
```

### .env.example

```env
# Claude API Key（服务端持有，下发给客户端）
CLAUDE_API_KEY=sk-ant-xxx
CLAUDE_BASE_URL=https://api.anthropic.com

# JWT 签名密钥
JWT_SECRET=your-random-secret-here
JWT_EXPIRE_DAYS=7

# 数据库
DATABASE_URL=postgresql+psycopg2://deepexcel:password@db:5432/deepexcel

# Redis
REDIS_URL=redis://redis:6379/0

# 订阅配置
TRIAL_DAYS=30
```

## 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| Web 框架 | FastAPI | 异步高性能，与 sidecar.py 同为 Python，生态一致 |
| ORM | SQLAlchemy 2.0 | Python 生态最成熟的 ORM |
| 迁移 | Alembic | SQLAlchemy 官方迁移工具 |
| 数据库 | PostgreSQL 16 | 成熟关系型数据库，JSON 字段支持好 |
| 缓存 | Redis 7 | JWT 黑名单、Key 缓存 |
| 密码哈希 | bcrypt (passlib) | 行业标准 |
| JWT | python-jose | FastAPI 生态常用 |
| 容器化 | Docker Compose | 一键启动，本地开发+生产部署一致 |

## 测试策略

1. **单元测试**：auth service、subscription service、key 下发逻辑
2. **API 集成测试**：FastAPI TestClient + 内存 SQLite，覆盖所有端点
3. **安全测试**：无 JWT 访问 /api/key 返回 401、过期订阅返回 403、重复注册返回 409
4. **手动测试**：Docker Compose 启动后用 curl 跑完整流程（注册→登录→拉 Key→上报日志）

## 不做的事（MVP 边界）

- ❌ 服务端代理 AI 流量（Agent loop 留本地）
- ❌ 用量统计 / token 计数 / 计费
- ❌ 支付集成（Stripe / 支付宝）
- ❌ 管理后台 UI
- ❌ 邮件验证（MVP 阶段不验证邮箱真实性）
- ❌ 第三方登录（OAuth）
- ❌ JWT 刷新机制（7 天过期后重新登录）
- ❌ 客户端改造（登录 UI + sidecar 启动流程）—— 后续 spec

## 后续迭代方向（不在本 spec）

1. **客户端改造**：登录 UI、sidecar 启动流程、日志上报
2. **用量统计**：客户端上报 token 用量，服务端记录+展示
3. **支付集成**：Stripe / 支付宝，订阅计划管理
4. **管理后台**：用户管理、订阅管理、日志查看
5. **邮件验证**：注册时发验证邮件
6. **JWT 刷新**：refresh token 机制
7. **Key 安全增强**：短期 token 代理模式（选项 B）
