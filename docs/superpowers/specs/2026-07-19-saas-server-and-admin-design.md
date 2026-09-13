# DeepExcel SaaS 服务端 + 管理后台

**日期：** 2026-07-19
**状态：** 设计已完成，待实现
**关联：**
- 替代并扩展 [2026-07-14-saas-server-mvp-design.md](2026-07-14-saas-server-mvp-design.md)（服务端 MVP 部分）
- 客户端改造仍参考 [2026-07-15-saas-client-integration-design.md](2026-07-15-saas-client-integration-design.md)
- 本 spec 新增管理后台 UI（前两个 spec 明确不包含的部分）

## 背景与动机

DeepExcel 当前是纯本地客户端，用户自带 API Key。要 SaaS 化需要：
1. **服务端**：注册/登录/Key 下发/日志上报（2026-07-14 已设计，本 spec 落地）
2. **管理后台**：运营人员管理用户、订阅、日志、API Key
3. **客户端改造**：登录 UI + sidecar 启动流程（2026-07-15 已设计，待落地）

本 spec 覆盖 **服务端 + 管理后台** 两个部分。客户端改造作为下一阶段工作。

## 核心架构决策

1. **Agent loop 留本地**——服务端只做"门卫 + 日志 + 管理"，不参与 AI 调用编排
2. **Phase 1 Key 下发**：服务端直接下发真实 Claude API Key（客户端内存缓存，不落盘）
3. **管理员独立体系**：单独的 `admins` 表，与用户表隔离，环境变量引导首次管理员
4. **技术栈**：FastAPI（后端）+ React/Vite（后台 UI）+ PostgreSQL + Redis + Docker Compose
5. **部署**：Docker Compose 一键启动（api / admin / db / redis 四个服务）

## 整体架构

```
┌─────────────────────────────────────────────────────────────┐
│ Excel 插件（客户端，本 spec 不涉及）                         │
│  LoginPanel → AuthClient → SessionManager → sidecar         │
└────────────────────────┬────────────────────────────────────┘
                         │ HTTPS
                         ▼
┌─────────────────────────────────────────────────────────────┐
│ Docker Compose                                              │
│                                                              │
│  ┌──────────────────┐  ┌──────────────────────────────────┐ │
│  │ api (FastAPI)    │  │ admin (React/Vite 静态 nginx)    │ │
│  │  :8000           │  │  :8080                           │ │
│  │                  │  │                                  │ │
│  │ /api/auth/*      │  │ / → index.html                   │ │
│  │ /api/key         │  │ 登录 / 用户管理 / 订阅管理 /     │ │
│  │ /api/subscription│  │ 日志查看 / API Key 管理          │ │
│  │ /api/logs        │  │                                  │ │
│  │ /admin/auth/*    │  │ 调用 /admin/api/* 拿数据         │ │
│  │ /admin/api/*     │  │                                  │ │
│  └────────┬─────────┘  └──────────────┬───────────────────┘ │
│           │                            │                      │
│           └──────────┬─────────────────┘                      │
│                      ▼                                        │
│  ┌──────────────────┐       ┌──────────────────┐             │
│  │ PostgreSQL 16    │       │ Redis 7          │             │
│  │ users/subs/logs/ │       │ JWT 黑名单/缓存  │             │
│  │ admins/api_keys  │       │                  │             │
│  └──────────────────┘       └──────────────────┘             │
└─────────────────────────────────────────────────────────────┘
```

**关键设计：**
- `api` 服务同时承载用户 API（`/api/*`）和管理 API（`/admin/api/*`）
- `admin` 服务是独立的 React 静态站点（nginx 托管），调用 `api` 服务的 `/admin/api/*`
- 两个服务共享同一份 FastAPI 代码，避免代码重复

**请求路径（admin 浏览器 → api）：**
- 浏览器访问 `http://localhost:8080/admin/login` 加载 React SPA
- SPA 内 axios 调 `http://localhost:8000/admin/api/*` 拿数据
- 开发模式：vite dev server 配 proxy 把 `/admin/api` 转发到 `http://localhost:8000`
- 生产模式：nginx 反向代理把 `/admin/api` 转发到 `http://api:8000`（容器内网）
- 浏览器只看 `http://localhost:8080`，不直接调 8000

**CORS：** 生产环境配置 `ADMIN_ORIGIN=https://admin.deepexcel.com`，FastAPI 只允许该 origin 跨域调 `/admin/api/*`

## 服务端设计

### 目录结构

```
server/
├── docker-compose.yml          # api + admin + db + redis
├── docker-compose.override.yml # 本地开发覆盖
├── .env.example
├── api/
│   ├── Dockerfile
│   ├── requirements.txt
│   ├── alembic.ini
│   ├── alembic/
│   │   └── versions/
│   ├── app/
│   │   ├── main.py             # FastAPI 入口，挂载用户/管理路由
│   │   ├── config.py           # 环境变量配置
│   │   ├── deps.py             # 依赖注入（当前用户/当前管理员）
│   │   ├── db/
│   │   │   ├── database.py     # SQLAlchemy 引擎/session
│   │   │   └── models.py       # User, Subscription, LogRecord, Admin, ApiKey
│   │   ├── auth/
│   │   │   ├── router.py       # /api/auth/register, /api/auth/login
│   │   │   ├── service.py      # 注册/登录逻辑
│   │   │   ├── jwt.py          # JWT 签发/校验（用户 + 管理员两套）
│   │   │   └── password.py     # bcrypt 包装
│   │   ├── key/
│   │   │   ├── router.py       # GET /api/key
│   │   │   └── service.py      # Key 选择 + 缓存
│   │   ├── subscription/
│   │   │   ├── router.py       # /api/subscription
│   │   │   └── service.py      # 订阅状态/到期/冻结
│   │   ├── logs/
│   │   │   ├── router.py       # POST /api/logs
│   │   │   └── models.py
│   │   └── admin/              # 管理后台专用 API
│   │       ├── router.py       # /admin/api/* 总路由
│   │       ├── auth.py         # /admin/auth/login, /admin/auth/me
│   │       ├── users.py        # 用户 CRUD + 禁用/启用
│   │       ├── subscriptions.py # 订阅调整
│   │       ├── logs.py         # 日志查询
│   │       └── api_keys.py     # API Key CRUD + 轮换
│   └── tests/
│       ├── conftest.py
│       ├── test_auth.py
│       ├── test_key.py
│       ├── test_admin_auth.py
│       ├── test_admin_users.py
│       └── test_admin_api_keys.py
├── admin/                      # React 管理后台
│   ├── Dockerfile
│   ├── package.json
│   ├── vite.config.ts
│   ├── nginx.conf
│   ├── src/
│   │   ├── main.tsx
│   │   ├── App.tsx
│   │   ├── api/                # 调用 /admin/api/*
│   │   ├── pages/
│   │   │   ├── Login.tsx
│   │   │   ├── Users.tsx
│   │   │   ├── Subscriptions.tsx
│   │   │   ├── Logs.tsx
│   │   │   └── ApiKeys.tsx
│   │   ├── components/
│   │   │   ├── Layout.tsx
│   │   │   ├── Table.tsx
│   │   │   └── ...
│   │   └── styles.css
│   └── ...
└── README.md
```

### 数据库模型

```python
# server/api/app/db/models.py

class User(Base):
    __tablename__ = "users"
    id = Column(UUID, primary_key=True, default=uuid4)
    email = Column(String(255), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)  # bcrypt
    status = Column(String(16), default="active")  # active | disabled
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login_at = Column(DateTime, nullable=True)
    subscription = relationship("Subscription", back_populates="user", uselist=False)
    logs = relationship("LogRecord", back_populates="user")

class Subscription(Base):
    __tablename__ = "subscriptions"
    id = Column(UUID, primary_key=True, default=uuid4)
    user_id = Column(UUID, ForeignKey("users.id"), unique=True, nullable=False)
    tier = Column(String(32), default="trial")  # trial | basic | pro
    status = Column(String(32), default="active")  # active | expired | frozen
    expires_at = Column(DateTime, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    user = relationship("User", back_populates="subscription")

class LogRecord(Base):
    __tablename__ = "logs"
    id = Column(UUID, primary_key=True, default=uuid4)
    user_id = Column(UUID, ForeignKey("users.id"), nullable=False, index=True)
    timestamp = Column(DateTime, nullable=False, index=True)
    level = Column(String(16), default="info")
    event = Column(String(64))
    message = Column(Text)
    context = Column(JSON)
    created_at = Column(DateTime, default=datetime.utcnow)
    user = relationship("User", back_populates="logs")

class Admin(Base):
    """管理员账号，独立于 users 表。环境变量引导首次创建。"""
    __tablename__ = "admins"
    id = Column(UUID, primary_key=True, default=uuid4)
    username = Column(String(64), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)  # bcrypt
    role = Column(String(16), default="admin")  # admin | super_admin
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login_at = Column(DateTime, nullable=True)

class ApiKey(Base):
    """服务端持有的 Claude API Key。可有多条，启用/禁用/轮换。"""
    __tablename__ = "api_keys"
    id = Column(UUID, primary_key=True, default=uuid4)
    label = Column(String(64), nullable=False)  # 备注名，如 "主力 - Anthropic 官方"
    provider = Column(String(32), default="anthropic")  # anthropic | deepseek | stepfun
    key_encrypted = Column(Text, nullable=False)  # Fernet 对称加密
    base_url = Column(String(255), default="https://api.anthropic.com")
    is_active = Column(Boolean, default=True, index=True)
    last_used_at = Column(DateTime, nullable=True)  # 用于"最久未用优先"负载均衡
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class SubscriptionChangeLog(Base):
    """订阅变更历史，供管理员审计和 UI 详情抽屉展示。"""
    __tablename__ = "subscription_change_logs"
    id = Column(UUID, primary_key=True, default=uuid4)
    user_id = Column(UUID, ForeignKey("users.id"), nullable=False, index=True)
    admin_id = Column(UUID, ForeignKey("admins.id"), nullable=True)  # null=系统自动变更
    change_type = Column(String(32), nullable=False)  # tier_change | status_change | extend | create
    old_value = Column(String(64), nullable=True)
    new_value = Column(String(64), nullable=True)
    extend_days = Column(Integer, nullable=True)
    note = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

class AdminAuditLog(Base):
    """管理员操作审计日志。所有 /admin/api/* 写操作都记一条。"""
    __tablename__ = "admin_audit_logs"
    id = Column(UUID, primary_key=True, default=uuid4)
    admin_id = Column(UUID, ForeignKey("admins.id"), nullable=False, index=True)
    action = Column(String(64), nullable=False)  # user.disable | user.delete | sub.update | apikey.create | apikey.rotate | ...
    target_type = Column(String(32), nullable=True)  # user | subscription | api_key
    target_id = Column(UUID, nullable=True)
    payload = Column(JSON)  # 请求体快照（脱敏后）
    ip = Column(String(45))  # IPv4/IPv6
    created_at = Column(DateTime, default=datetime.utcnow, index=True)
```

**索引：**
- `users.email` 唯一索引
- `subscriptions.user_id` 唯一索引
- `logs(user_id, timestamp)` 复合索引
- `admins.username` 唯一索引
- `api_keys.is_active` 普通索引（PostgreSQL 可优化为部分索引，开发用 SQLite 兼容）
- `admin_audit_logs(admin_id, created_at)` 复合索引
- `subscription_change_logs(user_id, created_at)` 复合索引

### 用户 API（沿用 2026-07-14 spec，无改动）

| 端点 | 方法 | 描述 |
|------|------|------|
| `/api/auth/register` | POST | 注册 |
| `/api/auth/login` | POST | 登录，返回 JWT |
| `/api/key` | GET | JWT + 订阅校验，下发 Claude API Key |
| `/api/subscription` | GET | 查订阅状态 |
| `/api/logs` | POST | 批量上报日志 |

详细请求/响应格式见 [2026-07-14 spec](2026-07-14-saas-server-mvp-design.md)。

### 管理 API（本 spec 新增）

#### 管理员认证

**首次引导：** 启动时读 `INITIAL_ADMIN_USERNAME` + `INITIAL_ADMIN_PASSWORD` 环境变量。若 `admins` 表为空，自动创建首个 super_admin 账号。后续启动跳过。

**并发安全：** 多 worker 启动时可能并发执行引导逻辑。靠 `admins.username` 唯一约束兜底——并发 INSERT 时只有一个成功，其他抛 IntegrityError 后 catch 并忽略（视为已创建）。

##### POST /admin/auth/login

**请求：**
```json
{ "username": "admin", "password": "securepassword" }
```

**响应 200：**
```json
{
  "admin_id": "uuid",
  "username": "admin",
  "role": "super_admin",
  "token": "jwt_admin_token",
  "expires_at": "2026-07-26T00:00:00Z"
}
```

**错误：** 401 用户名或密码错误

**JWT 区分：** 管理员 JWT payload 含 `type: "admin"` 字段，与用户 JWT（`type: "user"`）严格区分。中间件 `get_current_admin` 校验 `type == "admin"`，`get_current_user` 校验 `type == "user"`。两类 token 不可互用。

##### GET /admin/auth/me
返回当前登录管理员信息。需管理员 JWT。

##### POST /admin/auth/logout
将当前 JWT 加入 Redis 黑名单（TTL = 剩余有效期）。

#### 用户管理

##### GET /admin/api/users
分页查询用户列表。

**Query：** `page=1&size=20&search=email&status=active`

**响应：**
```json
{
  "items": [
    {
      "id": "uuid",
      "email": "user@example.com",
      "status": "active",
      "created_at": "...",
      "last_login_at": "...",
      "subscription": { "tier": "trial", "status": "active", "expires_at": "...", "days_remaining": 25 }
    }
  ],
  "total": 152,
  "page": 1,
  "size": 20
}
```

##### GET /admin/api/users/{user_id}
查单个用户详情，含订阅和最近 10 条日志摘要。

##### POST /admin/api/users/{user_id}/disable
禁用用户（`users.status = "disabled"`）。已禁用用户登录返回 403。

##### POST /admin/api/users/{user_id}/enable
启用用户。

##### DELETE /admin/api/users/{user_id}
硬删除用户（连带订阅、日志）。仅 super_admin 可调。返回 422 如果用户仍有活跃订阅（需先冻结）。

#### 订阅管理

##### GET /admin/api/subscriptions
分页查询所有订阅。可按 tier/status 筛选。

##### PATCH /admin/api/users/{user_id}/subscription
调整用户订阅。

**请求：**
```json
{
  "tier": "pro",
  "status": "active",
  "extend_days": 30
}
```

**逻辑：**
- `tier` 可改（trial → basic → pro）
- `status` 可改（active / expired / frozen）
- `extend_days` 正数延长 expires_at，负数缩短

**响应：** 更新后的订阅详情

#### 日志查看

##### GET /admin/api/logs
分页查询所有用户日志。

**Query：** `page=1&size=50&user_id=&event=&level=&start_date=&end_date=`

**响应：**
```json
{
  "items": [
    {
      "id": "uuid",
      "user_id": "uuid",
      "user_email": "user@example.com",
      "timestamp": "...",
      "level": "info",
      "event": "sidecar_start",
      "message": "sidecar started, model=claude-sonnet-4-5",
      "context": { "model": "claude-sonnet-4-5" }
    }
  ],
  "total": 1234,
  "page": 1,
  "size": 50
}
```

##### GET /admin/api/logs/stats
日志统计（按 event 聚合，按天分组）。用于后台首页概览。

**Query：** `days=7`

**响应：**
```json
{
  "by_day": [
    { "date": "2026-07-18", "total": 523, "errors": 12, "warnings": 38 }
  ],
  "by_event": [
    { "event": "tool_call", "count": 1234 },
    { "event": "sidecar_start", "count": 87 }
  ]
}
```

#### API Key 管理

##### GET /admin/api/api-keys
列出所有 API Key（不返回明文 key）。

**响应：**
```json
{
  "items": [
    {
      "id": "uuid",
      "label": "主力 - Anthropic 官方",
      "provider": "anthropic",
      "base_url": "https://api.anthropic.com",
      "is_active": true,
      "key_preview": "sk-ant-...api03",  // 前 7 + 后 4
      "created_at": "...",
      "updated_at": "..."
    }
  ]
}
```

##### POST /admin/api/api-keys
新增 API Key。

**请求：**
```json
{
  "label": "备用 - DeepSeek",
  "provider": "deepseek",
  "api_key": "sk-xxx",
  "base_url": "https://api.deepseek.com/anthropic"
}
```

**逻辑：** Fernet 对称加密后存 `key_encrypted`。

##### PATCH /admin/api/api-keys/{id}
更新 label / base_url / is_active。不能改 key 本身（要轮换请用 POST 新增 + 旧禁用）。

##### DELETE /admin/api/api-keys/{id}
硬删除。已禁用的 Key 才能删。

##### POST /admin/api/api-keys/{id}/rotate
轮换 Key（管理员粘贴新 key，旧 key 立即设为 inactive）。

**请求：** `{ "new_api_key": "sk-yyy" }`

#### 管理员审计日志

##### GET /admin/api/audit
分页查询管理员操作审计日志。仅 super_admin 可调。

**Query：** `page=1&size=50&admin_id=&action=&start_date=&end_date=`

**响应：**
```json
{
  "items": [
    {
      "id": "uuid",
      "admin_id": "uuid",
      "admin_username": "admin",
      "action": "user.disable",
      "target_type": "user",
      "target_id": "uuid",
      "payload": { "user_id": "uuid", "reason": "违规" },
      "ip": "1.2.3.4",
      "created_at": "..."
    }
  ],
  "total": 87,
  "page": 1,
  "size": 50
}
```

##### GET /admin/api/users/{user_id}/subscription-history
查指定用户的订阅变更历史（来自 `subscription_change_logs` 表）。

**响应：**
```json
{
  "items": [
    {
      "id": "uuid",
      "admin_username": "admin",
      "change_type": "extend",
      "old_value": null,
      "new_value": null,
      "extend_days": 30,
      "note": "用户反馈延期",
      "created_at": "..."
    }
  ]
}
```

### Key 下发逻辑（用户侧 `GET /api/key`）

当用户拉 Key 时，服务端从 `api_keys` 表选一条 `is_active=true` 的 Key：

1. **tier → provider 映射**（写死在 config.py，可调）：
   | 订阅 tier | 优先 provider | 兜底 provider |
   |----------|---------------|--------------|
   | trial | deepseek | anthropic |
   | basic | anthropic | deepseek |
   | pro | anthropic | (无兜底) |
2. **选择算法**：先按 tier 查对应 provider 的 active keys；无匹配则查兜底 provider；仍无则任意 active key
3. **负载均衡**：多条候选时，按 `last_used_at` 升序选最久未用的（NULL 视为最早），更新 `last_used_at = now()`
4. **解密**：Fernet 解密 `key_encrypted` 得到明文 key，返回给客户端
5. **缓存**：Redis 缓存 5 分钟（key=`user_key:{user_id}`），避免每次拉 Key 都查库解密；管理员禁用/删除 Key 时主动清缓存

### 安全考量

- **管理员密码**：bcrypt cost=12
- **管理员 JWT**：与用户 JWT 严格隔离（`type` 字段），有效期 24 小时（短于用户 7 天）
- **API Key 加密**：Fernet 对称加密（`FERNET_KEY` 环境变量），数据库被脱库不暴露 Key
- **管理员操作审计**：所有 `/admin/api/*` 写操作记 `admin_audit_logs` 表（admin_id, action, target, payload, timestamp）
- **HTTPS 强制**：生产环境 admin 和 api 都强制 HTTPS
- **Rate Limiting**：管理员登录限流（同 IP 每分钟 5 次），用户登录限流（同 IP 每分钟 10 次）
- **CORS**：admin 站点调 api 时需配置 CORS 白名单（只允许 admin 域名）
- **首次引导安全**：`INITIAL_ADMIN_PASSWORD` 环境变量只在 admins 表为空时读取，读取后立即从内存清除，不记日志

## 管理后台 UI 设计

### 技术栈

- React 18 + TypeScript + Vite
- React Router 6（SPA 路由）
- TanStack Query (React Query) v5（数据获取 + 缓存）
- Tailwind CSS（CodeX 简洁风格，6-10px 圆角，深灰+靛蓝配色）
- Lucide Icons（SVG 图标，符合用户偏好）
- Recharts（图表，用于日志统计可视化）
- npm build 后由 nginx 静态托管

### 页面结构

```
/admin/login        → 登录页（无侧边栏）
/admin/             → 仪表盘（首页）
  ├─ /users         → 用户管理
  ├─ /subscriptions → 订阅管理
  ├─ /logs          → 日志查看
  ├─ /api-keys      → API Key 管理
  └─ /audit         → 管理员操作审计（super_admin 可见）
```

### 全局布局

```
┌─────────────────────────────────────────────────────────┐
│  [Logo] DeepExcel Admin    [admin] [退出]               │ ← 顶部栏
├──────────┬──────────────────────────────────────────────┤
│ 仪表盘   │                                              │
│ 用户     │           主内容区（路由出口）                │
│ 订阅     │                                              │
│ 日志     │                                              │
│ API Key  │                                              │
│ 审计     │                                              │
└──────────┴──────────────────────────────────────────────┘
```

- 左侧侧边栏（200px 宽，图标 + 文字）
- 顶部栏显示当前管理员用户名 + 退出按钮
- 整体配色：浅色主题（白底 + 深灰文字 + 靛蓝强调）/ 暗色主题（深灰底 + 浅文字 + 天蓝强调）

### 各页面设计

#### 1. 登录页 `/admin/login`

简洁居中卡片：
- 用户名输入框
- 密码输入框
- 登录按钮（靛蓝）
- 错误提示（红色，下方）

登录成功后 JWT 存 localStorage，跳转 `/admin/`。

#### 2. 仪表盘 `/admin/`

四个数字卡片 + 一个折线图：
- 总用户数（含今日新增）
- 活跃订阅数
- 今日日志总数（含错误数）
- 当前启用 API Key 数

折线图：最近 7 天日志量趋势（按天聚合）。

#### 3. 用户管理 `/admin/users`

**表格列：**
- Email
- 状态（active 绿色 / disabled 红色）
- 订阅 tier + 到期天数
- 注册时间
- 最后登录
- 操作（查看详情 / 禁用 / 启用 / 删除）

**筛选：** 搜索框（按 email 模糊匹配）+ 状态下拉

**详情抽屉：** 点击某用户行从右侧滑出，显示：
- 用户基本信息
- 订阅详情（含历史调整记录）
- 最近 20 条日志

**操作确认：** 禁用/删除需二次确认（弹窗），删除操作仅 super_admin 可见。

#### 4. 订阅管理 `/admin/subscriptions`

**表格列：**
- Email
- Tier（trial/basic/pro）
- Status（active/expired/frozen）
- 到期时间
- 剩余天数
- 操作（编辑）

**编辑弹窗：**
- Tier 下拉
- Status 下拉
- 延长天数输入（可正可负）
- 保存按钮

#### 5. 日志查看 `/admin/logs`

**筛选栏：** 用户搜索 + event 下拉 + level 下拉 + 日期范围选择器

**表格列：**
- 时间
- 用户 Email
- Level（info 蓝 / warning 黄 / error 红）
- Event
- Message（截断，鼠标悬停显示完整）
- 详情按钮

**详情弹窗：** 显示完整 message 和 context JSON（pretty-print）。

#### 6. API Key 管理 `/admin/api-keys`

**表格列：**
- Label
- Provider
- Base URL
- Key 预览（`sk-ant-...api03`）
- 状态（启用/禁用）
- 创建时间
- 操作（编辑 / 轮换 / 删除）

**新增按钮：** 弹窗含 Label、Provider 下拉、API Key 输入、Base URL 输入。

**轮换操作：** 弹窗输入新 Key，确认后旧 Key 自动禁用。

**安全提示：** 新增/轮换时 Key 输入框旁提示"Key 将加密存储，无法找回明文，请妥善保管"。

### 前端约定

- 所有 API 调用走 `/admin/api/*` 前缀
- axios 拦截器自动带 JWT（`Authorization: Bearer xxx`）
- 401 响应自动跳登录页
- 表格用 TanStack Table v8（轻量，无 UI 框架依赖）
- 表单用 React Hook Form + Zod 校验
- 分页：固定底部，显示"共 X 条，第 Y/Z 页"
- 加载态：表格内容区骨架屏，按钮 loading spinner
- 错误处理：全局 toast（react-hot-toast），不阻塞操作

## 部署

### docker-compose.yml

```yaml
version: "3.8"
services:
  api:
    build: ./api
    ports:
      - "8000:8000"
    env_file: .env
    depends_on:
      - db
      - redis
    restart: unless-stopped

  admin:
    build: ./admin
    ports:
      - "8080:80"
    depends_on:
      - api
    restart: unless-stopped
    environment:
      - VITE_API_BASE=http://localhost:8000  # 构建时注入

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
# 数据库
DATABASE_URL=postgresql+psycopg2://deepexcel:password@db:5432/deepexcel
DB_PASSWORD=change-me

# Redis
REDIS_URL=redis://redis:6379/0

# JWT
JWT_SECRET=random-secret-for-user-jwt
JWT_EXPIRE_DAYS=7
ADMIN_JWT_SECRET=different-secret-for-admin-jwt
ADMIN_JWT_EXPIRE_HOURS=24

# API Key 加密（Fernet）
FERNET_KEY=generate-with-python-cryptography

# 首次管理员引导
INITIAL_ADMIN_USERNAME=admin
INITIAL_ADMIN_PASSWORD=change-me-immediately

# CORS
ADMIN_ORIGIN=http://localhost:8080
```

### 本地开发

```bash
# 启动所有服务
docker compose up -d

# 后端热重载（开发模式）
cd api
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

# 前端热重载（开发模式）
cd admin
npm install
npm run dev  # http://localhost:5173, proxy /admin/api → http://localhost:8000
```

### 数据库迁移

Alembic 管理 schema 变更：
```bash
cd api
alembic revision --autogenerate -m "initial schema"
alembic upgrade head
```

容器首次启动时自动执行 `alembic upgrade head`。

## 测试策略

### 后端测试

- **单元测试**：auth service、subscription service、key 选择逻辑、Fernet 加解密
- **API 集成测试**：FastAPI TestClient + 内存 SQLite，覆盖所有端点
- **安全测试**：
  - 用户 JWT 访问 `/admin/api/*` 返回 401
  - 管理员 JWT 访问 `/api/key` 返回 401
  - 禁用用户登录返回 403
  - 过期订阅拉 Key 返回 403
  - 首次引导：admins 表空时正确创建 super_admin，非空时跳过
  - 审计日志：所有写操作正确记录

### 前端测试

- **组件测试**：Vitest + Testing Library，覆盖表格、表单、弹窗
- **路由测试**：未登录跳转 login，登出跳 login
- **API mock**：MSW (Mock Service Worker) 模拟后端响应

### 端到端测试

- Docker Compose 启动后用 curl/Playwright 跑完整流程：
  - 管理员登录 → 创建 API Key → 启用
  - 用户注册 → 登录 → 拉 Key → 上报日志
  - 管理员查看用户/订阅/日志

## 不做的事（本 spec 边界）

- ❌ 客户端改造（LoginPanel + AuthClient 等，待 2026-07-15 spec 落地）
- ❌ 支付集成（Stripe / 支付宝）
- ❌ 邮件验证
- ❌ 用量统计 / token 计数 / 计费
- ❌ 多设备登录管理
- ❌ SSO/OAuth
- ❌ 管理员角色细分（只 admin / super_admin 两级）
- ❌ 国际化（后台纯中文）

## 实现顺序建议

1. **后端基础设施**：项目骨架 + Dockerfile + alembic + 数据库模型
2. **用户 API**：auth / key / subscription / logs（沿用 2026-07-14 spec）
3. **管理 API**：admin auth + users + subscriptions + logs + api-keys
4. **前端基础设施**：项目骨架 + 路由 + 全局布局 + 登录页
5. **前端业务页**：用户管理 → 订阅管理 → 日志查看 → API Key 管理 → 仪表盘
6. **Docker Compose 联调**：一键启动 + 端到端测试

## 后续迭代方向（不在本 spec）

1. **客户端改造落地**：按 2026-07-15 spec 实现 LoginPanel + AuthClient + SessionManager
2. **用量统计**：客户端上报 token 用量，后台展示
3. **支付集成**：Stripe / 支付宝
4. **Phase 2 代理模式**：客户端 base_url 指向服务端代理，Key 不再下发
5. **邮件验证 + 密码重置**
6. **多因素认证**（管理员）
7. **管理员角色细分**（运营 / 财务 / 超管）
