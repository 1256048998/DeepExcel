# DeepExcel SaaS 服务端 MVP 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 搭建 DeepExcel SaaS 服务端 MVP，提供用户注册/登录、订阅校验、API Key 下发、客户端日志上报四类接口，用 Docker Compose 一键部署。

**Architecture:** FastAPI 单体应用 + PostgreSQL + Redis。Agent loop 留客户端本地，服务端只做"门卫 + 日志"。Phase 1 直接下发真实 Claude API Key（Phase 2 演进为短期 token + 代理，不在本计划范围）。

**Tech Stack:** Python 3.11+、FastAPI、SQLAlchemy 2.0、Alembic、PostgreSQL 16、Redis 7、bcrypt（passlib）、python-jose（JWT）、Docker Compose、pytest

## Global Constraints

- 所有代码放 `server/` 目录（与 `src/` 平级，不污染现有客户端代码）
- Phase 1 范围：不做用量统计、不做支付、不做管理后台 UI、不做邮件验证、不做 OAuth、不做 JWT 刷新
- 数据库 ORM 用 SQLAlchemy 2.0 风格（`DeclarativeBase` + `Mapped` 类型注解）
- 测试用 SQLite 内存数据库 + FastAPI TestClient（生产用 PostgreSQL）
- 密码哈希 bcrypt cost=12
- JWT 有效期 7 天，签发时含 `user_id`、`tier`、`exp`
- 试用订阅 30 天，注册时自动创建
- API Key 仅存服务端环境变量，客户端拿到后只存内存（本计划不涉及客户端）
- JSON 字段用 `sqlalchemy.JSON`（兼容 SQLite 测试），不用 `postgresql.JSONB`
- 错误响应统一格式：`{"detail": "message"}`（FastAPI 默认）

---

### Task 1: 项目脚手架 + 配置文件

**Files:**
- Create: `server/requirements.txt`
- Create: `server/.env.example`
- Create: `server/Dockerfile`
- Create: `server/docker-compose.yml`
- Create: `server/app/__init__.py`（空文件）
- Create: `server/app/config.py`
- Create: `server/tests/__init__.py`（空文件）

**Interfaces:**
- Produces: `app.config.Settings` 类（Pydantic BaseSettings），后续所有模块通过 `get_settings()` 读取配置

- [ ] **Step 1: 创建 server/ 目录结构**

```bash
mkdir -p server/app server/tests
```

- [ ] **Step 2: 写 requirements.txt**

`server/requirements.txt`：

```
fastapi==0.115.0
uvicorn[standard]==0.30.6
sqlalchemy==2.0.35
alembic==1.13.3
psycopg2-binary==2.9.9
redis==5.0.8
passlib[bcrypt]==1.7.4
python-jose[cryptography]==3.3.0
pydantic==2.9.2
pydantic-settings==2.5.2
python-multipart==0.0.12
email-validator==2.2.0
pytest==8.3.3
httpx==0.27.2
```

- [ ] **Step 3: 写 .env.example**

`server/.env.example`：

```env
# Claude API Key（服务端持有，下发给客户端）
CLAUDE_API_KEY=sk-ant-xxx
CLAUDE_BASE_URL=https://api.anthropic.com

# JWT 签名密钥（生产环境必须改成随机长字符串）
JWT_SECRET=change-me-to-a-random-secret-at-least-32-chars
JWT_EXPIRE_DAYS=7
JWT_ALGORITHM=HS256

# 数据库（docker-compose 启动时用 db 主机名）
DATABASE_URL=postgresql+psycopg2://deepexcel:deepexcel_password@db:5432/deepexcel

# Redis
REDIS_URL=redis://redis:6379/0

# 订阅配置
TRIAL_DAYS=30

# CORS（客户端 Excel 插件不通过浏览器，但本地调试方便）
CORS_ORIGINS=http://localhost:3000,http://localhost:5173
```

- [ ] **Step 4: 写 Dockerfile**

`server/Dockerfile`：

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# 系统依赖（psycopg2 编译 + bcrypt）
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc libpq-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
```

- [ ] **Step 5: 写 docker-compose.yml**

`server/docker-compose.yml`：

```yaml
version: "3.8"

services:
  api:
    build: .
    ports:
      - "8000:8000"
    env_file: .env
    depends_on:
      db:
        condition: service_healthy
      redis:
        condition: service_started
    restart: unless-stopped

  db:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: deepexcel
      POSTGRES_USER: deepexcel
      POSTGRES_PASSWORD: deepexcel_password
    volumes:
      - pgdata:/var/lib/postgresql/data
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U deepexcel -d deepexcel"]
      interval: 5s
      timeout: 3s
      retries: 5
    restart: unless-stopped

  redis:
    image: redis:7-alpine
    restart: unless-stopped

volumes:
  pgdata:
```

- [ ] **Step 6: 写 app/__init__.py 和 tests/__init__.py**

两个文件都创建为空文件（Python 包标识）。

`server/app/__init__.py`：（空内容）

`server/tests/__init__.py`：（空内容）

- [ ] **Step 7: 写 app/config.py**

`server/app/config.py`：

```python
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """应用配置，从环境变量读取。"""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # Claude API
    claude_api_key: str
    claude_base_url: str = "https://api.anthropic.com"

    # JWT
    jwt_secret: str
    jwt_expire_days: int = 7
    jwt_algorithm: str = "HS256"

    # Database
    database_url: str

    # Redis
    redis_url: str = "redis://redis:6379/0"

    # Subscription
    trial_days: int = 30

    # CORS
    cors_origins: str = ""

    @property
    def cors_origins_list(self) -> list[str]:
        if not self.cors_origins:
            return []
        return [o.strip() for o in self.cors_origins.split(",") if o.strip()]


_settings: Settings | None = None


def get_settings() -> Settings:
    """单例 Settings（避免每次请求都重读 .env）。"""
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings
```

- [ ] **Step 8: 写配置加载测试**

`server/tests/test_config.py`：

```python
import os

from app.config import Settings


def test_settings_loads_from_env(monkeypatch):
    monkeypatch.setenv("CLAUDE_API_KEY", "sk-test-xxx")
    monkeypatch.setenv("JWT_SECRET", "a" * 32)
    monkeypatch.setenv("DATABASE_URL", "sqlite:///:memory:")
    monkeypatch.delenv("CLAUDE_BASE_URL", raising=False)
    monkeypatch.delenv("REDIS_URL", raising=False)
    monkeypatch.delenv("TRIAL_DAYS", raising=False)
    monkeypatch.delenv("CORS_ORIGINS", raising=False)

    settings = Settings()
    assert settings.claude_api_key == "sk-test-xxx"
    assert settings.claude_base_url == "https://api.anthropic.com"
    assert settings.jwt_expire_days == 7
    assert settings.trial_days == 30
    assert settings.cors_origins_list == []


def test_cors_origins_parses_csv(monkeypatch):
    monkeypatch.setenv("CLAUDE_API_KEY", "sk-test-xxx")
    monkeypatch.setenv("JWT_SECRET", "a" * 32)
    monkeypatch.setenv("DATABASE_URL", "sqlite:///:memory:")
    monkeypatch.setenv("CORS_ORIGINS", "http://localhost:3000, http://localhost:5173")

    settings = Settings()
    assert settings.cors_origins_list == [
        "http://localhost:3000",
        "http://localhost:5173",
    ]
```

- [ ] **Step 9: 安装依赖并跑测试**

Run: `cd server && pip install -r requirements.txt && python -m pytest tests/test_config.py -v`
Expected: 2 passed

- [ ] **Step 10: Commit**

```bash
git add server/
git commit -m "feat(server): scaffold FastAPI project with config and docker-compose"
```

---

### Task 2: 数据库连接 + ORM 模型 + Alembic 初始化

**Files:**
- Create: `server/app/db/__init__.py`（空文件）
- Create: `server/app/db/database.py`
- Create: `server/app/db/models.py`
- Create: `server/alembic.ini`
- Create: `server/alembic/env.py`
- Create: `server/alembic/script.py.mako`
- Create: `server/alembic/versions/`（空目录，放迁移文件）
- Test: `server/tests/test_models.py`

**Interfaces:**
- Produces: `app.db.database.Base`（DeclarativeBase）、`app.db.database.get_db()` 依赖、`app.db.database.engine`
- Produces: `app.db.models.User`、`app.db.models.Subscription`、`app.db.models.LogRecord`
- Consumes: `app.config.get_settings().database_url`

- [ ] **Step 1: 写 app/db/database.py**

`server/app/db/database.py`：

```python
from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from app.config import get_settings


class Base(DeclarativeBase):
    """所有 ORM 模型的基类。"""
    pass


_settings = get_settings()

# 测试环境用 SQLite，生产用 PostgreSQL
# SQLite 需要 check_same_thread=False 才能在 FastAPI 多线程用
if _settings.database_url.startswith("sqlite"):
    engine = create_engine(
        _settings.database_url,
        connect_args={"check_same_thread": False},
    )
else:
    engine = create_engine(_settings.database_url, pool_pre_ping=True)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


def get_db() -> Generator[Session, None, None]:
    """FastAPI 依赖：每请求一个 session，请求结束自动关闭。"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def create_all_tables() -> None:
    """测试用：直接建表（生产用 Alembic 迁移）。"""
    Base.metadata.create_all(bind=engine)
```

- [ ] **Step 2: 写 app/db/models.py**

`server/app/db/models.py`：

```python
import uuid
from datetime import datetime
from typing import Any

from sqlalchemy import DateTime, ForeignKey, String, Text
from sqlalchemy.dialects.postgresql import JSON, UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship
from sqlalchemy.types import JSON as SQLiteJSON

from app.db.database import Base


# 兼容 SQLite 测试：SQLite 不支持 postgresql.JSON，用通用 JSON
# SQLAlchemy 的 sqlalchemy.JSON 在 PostgreSQL 上也能用（只是没有 GIN 索引）
JSONType = SQLiteJSON


class User(Base):
    __tablename__ = "users"

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )
    email: Mapped[str] = mapped_column(String(255), unique=True, nullable=False, index=True)
    password_hash: Mapped[str] = mapped_column(String(255), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    subscription: Mapped["Subscription | None"] = relationship(
        "Subscription", back_populates="user", uselist=False, cascade="all, delete-orphan"
    )
    logs: Mapped[list["LogRecord"]] = relationship(
        "LogRecord", back_populates="user", cascade="all, delete-orphan"
    )


class Subscription(Base):
    __tablename__ = "subscriptions"

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )
    user_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("users.id"), unique=True, nullable=False, index=True
    )
    tier: Mapped[str] = mapped_column(String(32), default="trial")  # trial | basic | pro
    status: Mapped[str] = mapped_column(String(32), default="active")  # active | expired | frozen
    expires_at: Mapped[datetime] = mapped_column(DateTime, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.utcnow, onupdate=datetime.utcnow
    )

    user: Mapped["User"] = relationship("User", back_populates="subscription")


class LogRecord(Base):
    __tablename__ = "logs"

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )
    user_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("users.id"), nullable=False, index=True
    )
    timestamp: Mapped[datetime] = mapped_column(DateTime, nullable=False, index=True)
    level: Mapped[str] = mapped_column(String(16), default="info")  # info | warning | error
    event: Mapped[str | None] = mapped_column(String(64))
    message: Mapped[str | None] = mapped_column(Text)
    context: Mapped[dict[str, Any] | None] = mapped_column(JSONType)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    user: Mapped["User"] = relationship("User", back_populates="logs")
```

注意：`UUID(as_uuid=True)` 在 SQLite 下 SQLAlchemy 会回退为字符串存储（SQLAlchemy 2.0+ 原生支持），不影响测试。

- [ ] **Step 3: 写 conftest.py 提供 SQLite 测试 fixture**

`server/tests/conftest.py`：

```python
import os

# 必须在 import app 之前设置测试环境变量
os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "a" * 32)
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")
os.environ.setdefault("REDIS_URL", "redis://localhost:6379/0")

import pytest
from fastapi.testclient import TestClient
from sqlalchemy.orm import Session

from app.db.database import Base, SessionLocal, get_db
from app.main import app  # Task 8 创建


@pytest.fixture(scope="function")
def db_session() -> Session:
    """每个测试函数一个干净的内存数据库。"""
    Base.metadata.create_all(bind=app.dependency_overrides_get_engine())  # 见下面说明
    # 简化版：直接用全局 engine
    Base.metadata.drop_all()
    Base.metadata.create_all()
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all()


@pytest.fixture(scope="function")
def client(db_session: Session) -> TestClient:
    """FastAPI TestClient，覆盖 get_db 依赖用测试 session。"""

    def override_get_db():
        try:
            yield db_session
        finally:
            pass

    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()
```

简化版 conftest（去掉 `app.dependency_overrides_get_engine` 这行，因为没必要）：

`server/tests/conftest.py`（最终版）：

```python
import os

# 必须在 import app 之前设置测试环境变量
os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "a" * 32)
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")
os.environ.setdefault("REDIS_URL", "redis://localhost:6379/0")

import pytest
from fastapi.testclient import TestClient
from sqlalchemy.orm import Session

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app  # Task 8 创建


@pytest.fixture(scope="function")
def db_session() -> Session:
    """每个测试函数一个干净的内存数据库。"""
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session: Session) -> TestClient:
    """FastAPI TestClient，覆盖 get_db 依赖用测试 session。"""

    def override_get_db():
        try:
            yield db_session
        finally:
            pass

    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()
```

注意：此 conftest 依赖 `app.main`（Task 8 创建）。Task 2 测试先不依赖 client，只用 db_session。

- [ ] **Step 4: 写 test_models.py 验证模型可建表、可 CRUD**

`server/tests/test_models.py`：

```python
import uuid
from datetime import datetime, timedelta

from app.db.database import Base, SessionLocal, engine
from app.db.models import LogRecord, Subscription, User


def setup_function():
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)


def teardown_function():
    Base.metadata.drop_all(bind=engine)


def test_create_user_with_subscription():
    session = SessionLocal()
    try:
        user = User(
            email="test@example.com",
            password_hash="$2b$12$fakehash",
        )
        session.add(user)
        session.commit()
        session.refresh(user)

        sub = Subscription(
            user_id=user.id,
            tier="trial",
            status="active",
            expires_at=datetime.utcnow() + timedelta(days=30),
        )
        session.add(sub)
        session.commit()

        # 查回来
        loaded = session.query(User).filter_by(email="test@example.com").one()
        assert loaded.id is not None
        assert loaded.subscription.tier == "trial"
        assert loaded.subscription.status == "active"
    finally:
        session.close()


def test_create_log_record():
    session = SessionLocal()
    try:
        user = User(email="log@example.com", password_hash="$2b$12$fakehash")
        session.add(user)
        session.commit()
        session.refresh(user)

        log = LogRecord(
            user_id=user.id,
            timestamp=datetime.utcnow(),
            level="info",
            event="sidecar_start",
            message="sidecar started",
            context={"model": "claude-sonnet-4-5"},
        )
        session.add(log)
        session.commit()
        session.refresh(log)

        assert log.id is not None
        assert log.context == {"model": "claude-sonnet-4-5"}
        assert log.user.email == "log@example.com"
    finally:
        session.close()


def test_user_email_unique():
    session = SessionLocal()
    try:
        session.add(User(email="dup@example.com", password_hash="h1"))
        session.commit()
        session.add(User(email="dup@example.com", password_hash="h2"))
        # SQLite 不强制 unique？其实强制的，应该抛异常
        import pytest
        from sqlalchemy.exc import IntegrityError

        with pytest.raises(IntegrityError):
            session.commit()
        session.rollback()
    finally:
        session.close()
```

- [ ] **Step 5: 初始化 Alembic（先不跑命令，手动写最小配置）**

Alembic 用 `alembic init alembic` 会生成很多模板，但为了计划完整性，这里直接提供关键文件。

`server/alembic.ini`：

```ini
[alembic]
script_location = alembic
sqlalchemy.url = postgresql+psycopg2://deepexcel:deepexcel_password@db:5432/deepexcel

[loggers]
keys = root,sqlalchemy,alembic

[handlers]
keys = console

[formatters]
keys = generic

[logger_root]
level = WARN
handlers = console
qualname =

[logger_sqlalchemy]
level = WARN
handlers =
qualname = sqlalchemy.engine

[logger_alembic]
level = INFO
handlers =
qualname = alembic

[handler_console]
class = StreamHandler
args = (sys.stderr,)
level = NOTSET
formatter = generic

[formatter_generic]
format = %(levelname)-5.5s [%(name)s] %(message)s
datefmt = %H:%M:%S
```

`server/alembic/env.py`：

```python
from logging.config import fileConfig

from alembic import context
from sqlalchemy import engine_from_config, pool

from app.config import get_settings
from app.db.database import Base
from app.db import models  # noqa: F401 — 确保模型被注册到 metadata

config = context.config

if config.config_file_name is not None:
    fileConfig(config.config_file_name)

# 用应用配置覆盖 alembic.ini 里的 url
config.set_main_option("sqlalchemy.url", get_settings().database_url)

target_metadata = Base.metadata


def run_migrations_offline() -> None:
    url = config.get_main_option("sqlalchemy.url")
    context.configure(
        url=url,
        target_metadata=target_metadata,
        literal_binds=True,
        dialect_opts={"paramstyle": "named"},
    )
    with context.begin_transaction():
        context.run_migrations()


def run_migrations_online() -> None:
    connectable = engine_from_config(
        config.get_section(config.config_ini_section, {}),
        prefix="sqlalchemy.",
        poolclass=pool.NullPool,
    )
    with connectable.connect() as connection:
        context.configure(connection=connection, target_metadata=target_metadata)
        with context.begin_transaction():
            context.run_migrations()


if context.is_offline_mode():
    run_migrations_offline()
else:
    run_migrations_online()
```

`server/alembic/script.py.mako`：

```mako
"""${message}

Revision ID: ${up_revision}
Revises: ${down_revision | comma,n}
Create Date: ${create_date}

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
${imports if imports else ""}

revision: str = ${repr(up_revision)}
down_revision: Union[str, None] = ${repr(down_revision)}
branch_labels: Union[str, Sequence[str], None] = ${repr(branch_labels)}
depends_on: Union[str, Sequence[str], None] = ${repr(depends_on)}


def upgrade() -> None:
    ${upgrades if upgrades else "pass"}


def downgrade() -> None:
    ${downgrades if downgrades else "pass"}
```

创建空目录：`server/alembic/versions/`（放个 `.gitkeep` 文件）。

- [ ] **Step 6: 跑模型测试**

Run: `cd server && python -m pytest tests/test_models.py -v`
Expected: 3 passed

注意：此时 `app/main.py` 还没创建，conftest.py 会 import 失败。先把 conftest.py 的 `from app.main import app` 这行注释掉（Task 8 再放开），或者把 conftest 的 client fixture 暂时移除。

**简化方案：Task 2 阶段 conftest 只保留 db_session 相关，client fixture 在 Task 8 加。**

`server/tests/conftest.py`（Task 2 阶段版）：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "a" * 32)
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")
os.environ.setdefault("REDIS_URL", "redis://localhost:6379/0")
```

Task 8 再补充 fixture。

- [ ] **Step 7: 生成首个 Alembic 迁移（手动验证，不写测试）**

Run: `cd server && alembic revision --autogenerate -m "initial schema"`
Expected: 在 `alembic/versions/` 下生成一个 `xxxx_initial_schema.py` 文件，包含 users/subscriptions/logs 三张表的 `upgrade()` 和 `downgrade()`。

如果没装 alembic 或不想跑命令，跳过——Task 9 部署时会跑。

- [ ] **Step 8: Commit**

```bash
git add server/
git commit -m "feat(server): add database models and alembic setup"
```

---

### Task 3: JWT 工具模块

**Files:**
- Create: `server/app/auth/__init__.py`（空文件）
- Create: `server/app/auth/jwt.py`
- Test: `server/tests/test_jwt.py`

**Interfaces:**
- Consumes: `app.config.get_settings().jwt_secret`、`.jwt_algorithm`、`.jwt_expire_days`
- Produces: `app.auth.jwt.create_access_token(user_id: str, tier: str) -> str`
- Produces: `app.auth.jwt.decode_token(token: str) -> dict`（返回 `{"sub": user_id, "tier": "...", "exp": ...}`）

- [ ] **Step 1: 写 app/auth/jwt.py**

`server/app/auth/jwt.py`：

```python
from datetime import datetime, timedelta, timezone

from jose import JWTError, jwt

from app.config import get_settings


def create_access_token(user_id: str, tier: str) -> tuple[str, datetime]:
    """签发 JWT。返回 (token, expires_at)。"""
    settings = get_settings()
    expires_at = datetime.now(timezone.utc) + timedelta(days=settings.jwt_expire_days)
    payload = {
        "sub": user_id,
        "tier": tier,
        "exp": expires_at,
        "iat": datetime.now(timezone.utc),
    }
    token = jwt.encode(payload, settings.jwt_secret, algorithm=settings.jwt_algorithm)
    return token, expires_at


def decode_token(token: str) -> dict:
    """解码 JWT。返回 payload dict。失败抛 JWTError。"""
    settings = get_settings()
    payload = jwt.decode(token, settings.jwt_secret, algorithms=[settings.jwt_algorithm])
    return payload


class TokenInvalidError(Exception):
    """JWT 无效或过期。"""
    pass


def verify_token(token: str) -> dict:
    """校验 JWT，返回 payload。失败抛 TokenInvalidError。"""
    try:
        return decode_token(token)
    except JWTError as e:
        raise TokenInvalidError(str(e))
```

- [ ] **Step 2: 写 test_jwt.py**

`server/tests/test_jwt.py`：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")

from datetime import datetime, timezone

import pytest
from jose import jwt

from app.auth.jwt import TokenInvalidError, create_access_token, decode_token, verify_token
from app.config import get_settings


def test_create_and_decode_token():
    token, expires_at = create_access_token("user-uuid-123", "trial")
    assert token is not None
    assert isinstance(expires_at, datetime)

    payload = decode_token(token)
    assert payload["sub"] == "user-uuid-123"
    assert payload["tier"] == "trial"
    assert "exp" in payload


def test_verify_token_valid():
    token, _ = create_access_token("user-456", "basic")
    payload = verify_token(token)
    assert payload["sub"] == "user-456"
    assert payload["tier"] == "basic"


def test_verify_token_invalid_raises():
    with pytest.raises(TokenInvalidError):
        verify_token("invalid.token.here")


def test_token_expiry_in_future():
    token, expires_at = create_access_token("user-789", "trial")
    settings = get_settings()
    expected_expiry = datetime.now(timezone.utc) + timedelta(days=settings.jwt_expire_days)
    # 容忍 5 秒误差
    assert abs((expires_at - expected_expiry).total_seconds()) < 5
```

注意：上面 test 用了 `timedelta`，需要 import。修正：

```python
from datetime import datetime, timedelta, timezone
```

- [ ] **Step 3: 跑测试**

Run: `cd server && python -m pytest tests/test_jwt.py -v`
Expected: 4 passed

- [ ] **Step 4: Commit**

```bash
git add server/app/auth/ server/tests/test_jwt.py
git commit -m "feat(server): add JWT create/decode/verify utilities"
```

---

### Task 4: Auth 模块 — 注册/登录（service + router + 依赖）

**Files:**
- Create: `server/app/auth/dependencies.py` — FastAPI 依赖：从请求头提取 JWT，返回 user_id
- Create: `server/app/auth/service.py` — 注册/登录业务逻辑
- Create: `server/app/auth/router.py` — `/api/auth/register`、`/api/auth/login`
- Create: `server/app/auth/schemas.py` — Pydantic 请求/响应模型
- Test: `server/tests/test_auth.py`

**Interfaces:**
- Consumes: `app.db.models.User`、`app.db.models.Subscription`、`app.auth.jwt.create_access_token`、`app.config.get_settings().trial_days`
- Produces: `app.auth.dependencies.get_current_user_id` FastAPI 依赖（后续所有受保护端点用）
- Produces: `app.auth.router.router`（FastAPI APIRouter，挂载到 `/api/auth`）

- [ ] **Step 1: 写 app/auth/schemas.py**

`server/app/auth/schemas.py`：

```python
from datetime import datetime

from pydantic import BaseModel, EmailStr, Field


class RegisterRequest(BaseModel):
    email: EmailStr
    password: str = Field(min_length=8, max_length=128)


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class AuthResponse(BaseModel):
    user_id: str
    email: str
    token: str
    expires_at: datetime
```

- [ ] **Step 2: 写 app/auth/service.py**

`server/app/auth/service.py`：

```python
from datetime import datetime, timedelta, timezone

from sqlalchemy.orm import Session

from app.config import get_settings
from app.db.models import Subscription, User


class AuthError(Exception):
    """认证业务错误。"""
    def __init__(self, message: str, status_code: int = 400):
        self.message = message
        self.status_code = status_code
        super().__init__(message)


def register_user(db: Session, email: str, password: str) -> tuple[User, str, datetime]:
    """注册新用户。返回 (user, token, expires_at)。
    
    - email 已存在 → AuthError(409)
    - 密码 < 8 字符 → AuthError(400)（Pydantic 已校验，这里是兜底）
    """
    existing = db.query(User).filter_by(email=email).first()
    if existing:
        raise AuthError("邮箱已注册", status_code=409)

    from passlib.context import CryptContext
    pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto", bcrypt__rounds=12)
    password_hash = pwd_context.hash(password)

    user = User(email=email, password_hash=password_hash)
    db.add(user)
    db.flush()  # 拿 user.id

    settings = get_settings()
    sub = Subscription(
        user_id=user.id,
        tier="trial",
        status="active",
        expires_at=datetime.now(timezone.utc) + timedelta(days=settings.trial_days),
    )
    db.add(sub)
    db.commit()
    db.refresh(user)

    from app.auth.jwt import create_access_token
    token, expires_at = create_access_token(str(user.id), "trial")
    return user, token, expires_at


def login_user(db: Session, email: str, password: str) -> tuple[User, str, datetime]:
    """登录。返回 (user, token, expires_at)。
    
    - 用户不存在或密码错误 → AuthError(401)
    """
    user = db.query(User).filter_by(email=email).first()
    if not user:
        raise AuthError("邮箱或密码错误", status_code=401)

    from passlib.context import CryptContext
    pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
    if not pwd_context.verify(password, user.password_hash):
        raise AuthError("邮箱或密码错误", status_code=401)

    # 查订阅 tier 用于 JWT
    tier = user.subscription.tier if user.subscription else "trial"

    from app.auth.jwt import create_access_token
    token, expires_at = create_access_token(str(user.id), tier)
    return user, token, expires_at
```

- [ ] **Step 3: 写 app/auth/dependencies.py**

`server/app/auth/dependencies.py`：

```python
from fastapi import Depends, Header, HTTPException
from sqlalchemy.orm import Session

from app.auth.jwt import TokenInvalidError, verify_token
from app.db.database import get_db


def get_current_user_id(
    authorization: str | None = Header(None),
) -> str:
    """FastAPI 依赖：从 Authorization: Bearer <token> 提取 user_id。
    
    失败抛 401。
    """
    if not authorization:
        raise HTTPException(status_code=401, detail="缺少 Authorization 头")
    if not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Authorization 格式错误，应为 Bearer <token>")
    token = authorization.removeprefix("Bearer ").strip()
    try:
        payload = verify_token(token)
    except TokenInvalidError as e:
        raise HTTPException(status_code=401, detail=f"JWT 无效或过期: {e}")
    user_id = payload.get("sub")
    if not user_id:
        raise HTTPException(status_code=401, detail="JWT 缺少 sub 字段")
    return user_id
```

- [ ] **Step 4: 写 app/auth/router.py**

`server/app/auth/router.py`：

```python
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.auth.dependencies import get_current_user_id
from app.auth.schemas import AuthResponse, LoginRequest, RegisterRequest
from app.auth.service import AuthError, login_user, register_user
from app.db.database import get_db

router = APIRouter(prefix="/api/auth", tags=["auth"])


@router.post("/register", response_model=AuthResponse, status_code=201)
def register(req: RegisterRequest, db: Session = Depends(get_db)):
    try:
        user, token, expires_at = register_user(db, req.email, req.password)
    except AuthError as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)
    return AuthResponse(
        user_id=str(user.id),
        email=user.email,
        token=token,
        expires_at=expires_at,
    )


@router.post("/login", response_model=AuthResponse)
def login(req: LoginRequest, db: Session = Depends(get_db)):
    try:
        user, token, expires_at = login_user(db, req.email, req.password)
    except AuthError as e:
        raise HTTPException(status_code=e.status_code, detail=e.message)
    return AuthResponse(
        user_id=str(user.id),
        email=user.email,
        token=token,
        expires_at=expires_at,
    )


@router.get("/me")
def me(user_id: str = Depends(get_current_user_id)):
    """测试用：验证 JWT 是否有效。"""
    return {"user_id": user_id}
```

- [ ] **Step 5: 写 test_auth.py**

`server/tests/test_auth.py`：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")

import pytest
from fastapi.testclient import TestClient
from sqlalchemy.orm import Session

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app  # Task 8 创建


@pytest.fixture(scope="function")
def db_session():
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session):
    def override():
        yield db_session

    app.dependency_overrides[get_db] = override
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()


def test_register_success(client):
    resp = client.post("/api/auth/register", json={
        "email": "new@example.com",
        "password": "securepassword123",
    })
    assert resp.status_code == 201
    data = resp.json()
    assert data["email"] == "new@example.com"
    assert data["token"]
    assert data["user_id"]
    assert "expires_at" in data


def test_register_duplicate_email(client):
    client.post("/api/auth/register", json={
        "email": "dup@example.com",
        "password": "securepassword123",
    })
    resp = client.post("/api/auth/register", json={
        "email": "dup@example.com",
        "password": "anotherpassword",
    })
    assert resp.status_code == 409


def test_register_short_password(client):
    resp = client.post("/api/auth/register", json={
        "email": "short@example.com",
        "password": "1234567",  # 7 字符
    })
    assert resp.status_code == 422  # Pydantic 校验失败


def test_register_invalid_email(client):
    resp = client.post("/api/auth/register", json={
        "email": "not-an-email",
        "password": "securepassword123",
    })
    assert resp.status_code == 422


def test_login_success(client):
    client.post("/api/auth/register", json={
        "email": "login@example.com",
        "password": "securepassword123",
    })
    resp = client.post("/api/auth/login", json={
        "email": "login@example.com",
        "password": "securepassword123",
    })
    assert resp.status_code == 200
    assert resp.json()["token"]


def test_login_wrong_password(client):
    client.post("/api/auth/register", json={
        "email": "wrong@example.com",
        "password": "securepassword123",
    })
    resp = client.post("/api/auth/login", json={
        "email": "wrong@example.com",
        "password": "wrongpassword",
    })
    assert resp.status_code == 401


def test_login_nonexistent_user(client):
    resp = client.post("/api/auth/login", json={
        "email": "nobody@example.com",
        "password": "whatever",
    })
    assert resp.status_code == 401


def test_me_with_valid_token(client):
    reg = client.post("/api/auth/register", json={
        "email": "me@example.com",
        "password": "securepassword123",
    })
    token = reg.json()["token"]
    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200
    assert resp.json()["user_id"] == reg.json()["user_id"]


def test_me_without_token(client):
    resp = client.get("/api/auth/me")
    assert resp.status_code == 401


def test_me_with_invalid_token(client):
    resp = client.get("/api/auth/me", headers={"Authorization": "Bearer invalid.token.here"})
    assert resp.status_code == 401
```

注意：此测试依赖 `app.main`（Task 8 创建）。Task 4 阶段可以先创建一个最小 `app/main.py` 只挂载 auth router，Task 8 再补全。

**简化：Task 4 之前先创建最小 app/main.py，Task 8 再扩展。**

在 Step 5 之前先做 Step 4.5（见下）。

- [ ] **Step 4.5: 创建最小 app/main.py（仅挂 auth router）**

`server/app/main.py`：

```python
from fastapi import FastAPI

from app.auth.router import router as auth_router

app = FastAPI(title="DeepExcel Server", version="0.1.0")

app.include_router(auth_router)


@app.get("/health")
def health():
    return {"status": "ok"}
```

- [ ] **Step 6: 跑 auth 测试**

Run: `cd server && python -m pytest tests/test_auth.py -v`
Expected: 10 passed

注意：bcrypt cost=12 在测试里会比较慢（每次注册/登录约 250ms）。如果太慢，可在 conftest 里 monkeypatch 把 rounds 降到 4，但本计划不强制要求。

- [ ] **Step 7: Commit**

```bash
git add server/app/auth/ server/app/main.py server/tests/test_auth.py
git commit -m "feat(server): add auth module (register, login, JWT dependency)"
```

---

### Task 5: Subscription 模块

**Files:**
- Create: `server/app/subscription/__init__.py`（空文件）
- Create: `server/app/subscription/schemas.py`
- Create: `server/app/subscription/service.py`
- Create: `server/app/subscription/router.py`
- Test: `server/tests/test_subscription.py`

**Interfaces:**
- Consumes: `app.auth.dependencies.get_current_user_id`、`app.db.models.Subscription`
- Produces: `app.subscription.router.router`（挂载到 `/api/subscription`）
- Produces: `app.subscription.service.get_subscription_status(db, user_id) -> dict`（返回 `{tier, status, expires_at, days_remaining}`）
- Produces: `app.subscription.service.is_subscription_active(db, user_id) -> bool`（key 下发模块用）

- [ ] **Step 1: 写 app/subscription/schemas.py**

`server/app/subscription/schemas.py`：

```python
from datetime import datetime

from pydantic import BaseModel


class SubscriptionResponse(BaseModel):
    tier: str
    status: str
    expires_at: datetime
    days_remaining: int
```

- [ ] **Step 2: 写 app/subscription/service.py**

`server/app/subscription/service.py`：

```python
from datetime import datetime, timezone

from sqlalchemy.orm import Session

from app.db.models import Subscription


def get_subscription_status(db: Session, user_id: str) -> dict | None:
    """查订阅状态。返回 {tier, status, expires_at, days_remaining}。
    
    订阅不存在返回 None。
    """
    sub = db.query(Subscription).filter_by(user_id=user_id).first()
    if not sub:
        return None

    now = datetime.now(timezone.utc)
    # sub.expires_at 可能是 naive datetime（SQLite）或 aware（PostgreSQL）
    expires_at = sub.expires_at
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)

    days_remaining = max(0, (expires_at - now).days)

    # 如果数据库里 status=active 但已过期，自动标记 expired（不写回 DB，仅返回）
    status = sub.status
    if status == "active" and expires_at < now:
        status = "expired"

    return {
        "tier": sub.tier,
        "status": status,
        "expires_at": expires_at,
        "days_remaining": days_remaining,
    }


def is_subscription_active(db: Session, user_id: str) -> bool:
    """订阅是否有效（key 下发模块用）。"""
    status = get_subscription_status(db, user_id)
    if not status:
        return False
    return status["status"] == "active" and status["days_remaining"] > 0
```

- [ ] **Step 3: 写 app/subscription/router.py**

`server/app/subscription/router.py`：

```python
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.auth.dependencies import get_current_user_id
from app.db.database import get_db
from app.subscription.schemas import SubscriptionResponse
from app.subscription.service import get_subscription_status

router = APIRouter(prefix="/api/subscription", tags=["subscription"])


@router.get("", response_model=SubscriptionResponse)
def get_subscription(
    user_id: str = Depends(get_current_user_id),
    db: Session = Depends(get_db),
):
    status = get_subscription_status(db, user_id)
    if not status:
        raise HTTPException(status_code=404, detail="订阅记录不存在")
    return SubscriptionResponse(**status)
```

- [ ] **Step 4: 在 app/main.py 中挂载 subscription router**

修改 `server/app/main.py`，添加：

```python
from app.subscription.router import router as subscription_router

app.include_router(subscription_router)
```

完整版：

```python
from fastapi import FastAPI

from app.auth.router import router as auth_router
from app.subscription.router import router as subscription_router

app = FastAPI(title="DeepExcel Server", version="0.1.0")

app.include_router(auth_router)
app.include_router(subscription_router)


@app.get("/health")
def health():
    return {"status": "ok"}
```

- [ ] **Step 5: 写 test_subscription.py**

`server/tests/test_subscription.py`：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")

import pytest
from fastapi.testclient import TestClient

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app


@pytest.fixture(scope="function")
def db_session():
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session):
    def override():
        yield db_session
    app.dependency_overrides[get_db] = override
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()


def _register_and_get_token(client, email="sub@example.com"):
    resp = client.post("/api/auth/register", json={
        "email": email,
        "password": "securepassword123",
    })
    return resp.json()["token"]


def test_get_subscription_active(client):
    token = _register_and_get_token(client)
    resp = client.get("/api/subscription", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200
    data = resp.json()
    assert data["tier"] == "trial"
    assert data["status"] == "active"
    assert data["days_remaining"] > 0


def test_get_subscription_no_token(client):
    resp = client.get("/api/subscription")
    assert resp.status_code == 401


def test_subscription_days_remaining_decreases(client, db_session):
    """手动改 expires_at 验证 days_remaining 计算。"""
    token = _register_and_get_token(client)
    
    # 直接改 DB，把过期时间设为 5 天后
    from app.db.models import Subscription
    from datetime import datetime, timedelta, timezone
    sub = db_session.query(Subscription).first()
    sub.expires_at = datetime.now(timezone.utc) + timedelta(days=5)
    db_session.commit()

    resp = client.get("/api/subscription", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200
    assert resp.json()["days_remaining"] == 5


def test_subscription_expired(client, db_session):
    token = _register_and_get_token(client)
    
    from app.db.models import Subscription
    from datetime import datetime, timedelta, timezone
    sub = db_session.query(Subscription).first()
    sub.expires_at = datetime.now(timezone.utc) - timedelta(days=1)
    db_session.commit()

    resp = client.get("/api/subscription", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200
    assert resp.json()["status"] == "expired"
    assert resp.json()["days_remaining"] == 0
```

- [ ] **Step 6: 跑测试**

Run: `cd server && python -m pytest tests/test_subscription.py -v`
Expected: 4 passed

- [ ] **Step 7: Commit**

```bash
git add server/app/subscription/ server/app/main.py server/tests/test_subscription.py
git commit -m "feat(server): add subscription module with status check"
```

---

### Task 6: Key 下发模块

**Files:**
- Create: `server/app/key/__init__.py`（空文件）
- Create: `server/app/key/schemas.py`
- Create: `server/app/key/router.py`
- Test: `server/tests/test_key.py`

**Interfaces:**
- Consumes: `app.auth.dependencies.get_current_user_id`、`app.subscription.service.is_subscription_active`、`app.config.get_settings().claude_api_key`、`.claude_base_url`
- Produces: `app.key.router.router`（挂载到 `/api/key`）

- [ ] **Step 1: 写 app/key/schemas.py**

`server/app/key/schemas.py`：

```python
from pydantic import BaseModel


class KeyResponse(BaseModel):
    api_key: str
    provider: str
    base_url: str
```

- [ ] **Step 2: 写 app/key/router.py**

`server/app/key/router.py`：

```python
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.auth.dependencies import get_current_user_id
from app.config import get_settings
from app.db.database import get_db
from app.key.schemas import KeyResponse
from app.subscription.service import is_subscription_active

router = APIRouter(prefix="/api/key", tags=["key"])


@router.get("", response_model=KeyResponse)
def get_api_key(
    user_id: str = Depends(get_current_user_id),
    db: Session = Depends(get_db),
):
    if not is_subscription_active(db, user_id):
        raise HTTPException(status_code=403, detail="订阅已过期或冻结")

    settings = get_settings()
    return KeyResponse(
        api_key=settings.claude_api_key,
        provider="anthropic",
        base_url=settings.claude_base_url,
    )
```

- [ ] **Step 3: 在 app/main.py 中挂载 key router**

修改 `server/app/main.py`，添加：

```python
from app.key.router import router as key_router

app.include_router(key_router)
```

完整版：

```python
from fastapi import FastAPI

from app.auth.router import router as auth_router
from app.key.router import router as key_router
from app.subscription.router import router as subscription_router

app = FastAPI(title="DeepExcel Server", version="0.1.0")

app.include_router(auth_router)
app.include_router(subscription_router)
app.include_router(key_router)


@app.get("/health")
def health():
    return {"status": "ok"}
```

- [ ] **Step 4: 写 test_key.py**

`server/tests/test_key.py`：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")

import pytest
from fastapi.testclient import TestClient

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app


@pytest.fixture(scope="function")
def db_session():
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session):
    def override():
        yield db_session
    app.dependency_overrides[get_db] = override
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()


def _register_and_get_token(client, email="key@example.com"):
    resp = client.post("/api/auth/register", json={
        "email": email,
        "password": "securepassword123",
    })
    return resp.json()["token"]


def test_get_key_success(client):
    token = _register_and_get_token(client)
    resp = client.get("/api/key", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200
    data = resp.json()
    assert data["api_key"] == "sk-test-xxx"
    assert data["provider"] == "anthropic"
    assert data["base_url"] == "https://api.anthropic.com"


def test_get_key_no_token(client):
    resp = client.get("/api/key")
    assert resp.status_code == 401


def test_get_key_expired_subscription(client, db_session):
    token = _register_and_get_token(client)
    
    # 把订阅设为已过期
    from app.db.models import Subscription
    from datetime import datetime, timedelta, timezone
    sub = db_session.query(Subscription).first()
    sub.expires_at = datetime.now(timezone.utc) - timedelta(days=1)
    db_session.commit()

    resp = client.get("/api/key", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 403


def test_get_key_frozen_subscription(client, db_session):
    token = _register_and_get_token(client)
    
    from app.db.models import Subscription
    sub = db_session.query(Subscription).first()
    sub.status = "frozen"
    db_session.commit()

    resp = client.get("/api/key", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 403
```

- [ ] **Step 5: 跑测试**

Run: `cd server && python -m pytest tests/test_key.py -v`
Expected: 4 passed

- [ ] **Step 6: Commit**

```bash
git add server/app/key/ server/app/main.py server/tests/test_key.py
git commit -m "feat(server): add API key distribution endpoint with subscription check"
```

---

### Task 7: Logs 模块（批量日志上报）

**Files:**
- Create: `server/app/logs/__init__.py`（空文件）
- Create: `server/app/logs/schemas.py`
- Create: `server/app/logs/router.py`
- Test: `server/tests/test_logs.py`

**Interfaces:**
- Consumes: `app.auth.dependencies.get_current_user_id`、`app.db.models.LogRecord`
- Produces: `app.logs.router.router`（挂载到 `/api/logs`）

- [ ] **Step 1: 写 app/logs/schemas.py**

`server/app/logs/schemas.py`：

```python
from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class LogEntry(BaseModel):
    timestamp: datetime
    level: str = Field(default="info", pattern="^(info|warning|error)$")
    event: str | None = None
    message: str | None = None
    context: dict[str, Any] | None = None


class LogBatchRequest(BaseModel):
    logs: list[LogEntry] = Field(min_length=1, max_length=100)


class LogBatchResponse(BaseModel):
    accepted: int
```

- [ ] **Step 2: 写 app/logs/router.py**

`server/app/logs/router.py`：

```python
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.auth.dependencies import get_current_user_id
from app.db.database import get_db
from app.db.models import LogRecord
from app.logs.schemas import LogBatchRequest, LogBatchResponse

router = APIRouter(prefix="/api/logs", tags=["logs"])


@router.post("", response_model=LogBatchResponse)
def upload_logs(
    req: LogBatchRequest,
    user_id: str = Depends(get_current_user_id),
    db: Session = Depends(get_db),
):
    """批量上报客户端日志。单次最多 100 条。"""
    import uuid

    records = []
    for entry in req.logs:
        records.append(LogRecord(
            id=uuid.uuid4(),
            user_id=uuid.UUID(user_id),
            timestamp=entry.timestamp,
            level=entry.level,
            event=entry.event,
            message=entry.message,
            context=entry.context,
        ))
    db.add_all(records)
    db.commit()

    return LogBatchResponse(accepted=len(records))
```

- [ ] **Step 3: 在 app/main.py 中挂载 logs router**

修改 `server/app/main.py`，添加：

```python
from app.logs.router import router as logs_router

app.include_router(logs_router)
```

完整版：

```python
from fastapi import FastAPI

from app.auth.router import router as auth_router
from app.key.router import router as key_router
from app.logs.router import router as logs_router
from app.subscription.router import router as subscription_router

app = FastAPI(title="DeepExcel Server", version="0.1.0")

app.include_router(auth_router)
app.include_router(subscription_router)
app.include_router(key_router)
app.include_router(logs_router)


@app.get("/health")
def health():
    return {"status": "ok"}
```

- [ ] **Step 4: 写 test_logs.py**

`server/tests/test_logs.py`：

```python
import os

os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")

import pytest
from fastapi.testclient import TestClient

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app


@pytest.fixture(scope="function")
def db_session():
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session):
    def override():
        yield db_session
    app.dependency_overrides[get_db] = override
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()


def _register_and_get_token(client, email="logs@example.com"):
    resp = client.post("/api/auth/register", json={
        "email": email,
        "password": "securepassword123",
    })
    return resp.json()["token"]


def test_upload_logs_success(client):
    token = _register_and_get_token(client)
    resp = client.post("/api/logs", 
        headers={"Authorization": f"Bearer {token}"},
        json={"logs": [
            {
                "timestamp": "2026-07-14T10:30:00Z",
                "level": "info",
                "event": "sidecar_start",
                "message": "sidecar started",
                "context": {"model": "claude-sonnet-4-5"},
            },
            {
                "timestamp": "2026-07-14T10:31:00Z",
                "level": "warning",
                "event": "tool_call",
                "message": "read_range returned 500 rows",
                "context": {"tool": "read_range", "rows": 500},
            },
        ]},
    )
    assert resp.status_code == 200
    assert resp.json()["accepted"] == 2


def test_upload_logs_no_token(client):
    resp = client.post("/api/logs", json={"logs": [
        {"timestamp": "2026-07-14T10:30:00Z", "level": "info"},
    ]})
    assert resp.status_code == 401


def test_upload_logs_empty_batch_rejected(client):
    token = _register_and_get_token(client)
    resp = client.post("/api/logs",
        headers={"Authorization": f"Bearer {token}"},
        json={"logs": []},
    )
    assert resp.status_code == 422  # min_length=1


def test_upload_logs_invalid_level_rejected(client):
    token = _register_and_get_token(client)
    resp = client.post("/api/logs",
        headers={"Authorization": f"Bearer {token}"},
        json={"logs": [
            {"timestamp": "2026-07-14T10:30:00Z", "level": "debug"},  # 不允许
        ]},
    )
    assert resp.status_code == 422


def test_upload_logs_persisted_to_db(client, db_session):
    token = _register_and_get_token(client)
    client.post("/api/logs",
        headers={"Authorization": f"Bearer {token}"},
        json={"logs": [
            {"timestamp": "2026-07-14T10:30:00Z", "level": "error", "event": "crash", "message": "boom"},
        ]},
    )
    
    from app.db.models import LogRecord
    logs = db_session.query(LogRecord).all()
    assert len(logs) == 1
    assert logs[0].level == "error"
    assert logs[0].event == "crash"
    assert logs[0].message == "boom"
```

- [ ] **Step 5: 跑测试**

Run: `cd server && python -m pytest tests/test_logs.py -v`
Expected: 5 passed

- [ ] **Step 6: Commit**

```bash
git add server/app/logs/ server/app/main.py server/tests/test_logs.py
git commit -m "feat(server): add batch log upload endpoint"
```

---

### Task 8: 集成测试 + CORS + main.py 完善

**Files:**
- Modify: `server/app/main.py` — 添加 CORS 中间件、异常处理器
- Create: `server/tests/test_integration.py` — 完整流程测试
- Update: `server/tests/conftest.py` — 补全 client fixture

**Interfaces:**
- Produces: 完整的 FastAPI app，所有 router 挂载、CORS 启用

- [ ] **Step 1: 完善 app/main.py（加 CORS）**

`server/app/main.py`：

```python
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.auth.router import router as auth_router
from app.config import get_settings
from app.key.router import router as key_router
from app.logs.router import router as logs_router
from app.subscription.router import router as subscription_router

settings = get_settings()

app = FastAPI(title="DeepExcel Server", version="0.1.0")

# CORS（Excel 插件不通过浏览器，但本地调试方便）
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list or ["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(auth_router)
app.include_router(subscription_router)
app.include_router(key_router)
app.include_router(logs_router)


@app.get("/health")
def health():
    return {"status": "ok", "version": "0.1.0"}
```

- [ ] **Step 2: 完善 conftest.py（统一 fixture）**

`server/tests/conftest.py`：

```python
import os

# 必须在 import app 之前设置测试环境变量
os.environ.setdefault("CLAUDE_API_KEY", "sk-test-xxx")
os.environ.setdefault("JWT_SECRET", "test-secret-32-chars-minimum-length!")
os.environ.setdefault("DATABASE_URL", "sqlite:///:memory:")
os.environ.setdefault("REDIS_URL", "redis://localhost:6379/0")

import pytest
from fastapi.testclient import TestClient
from sqlalchemy.orm import Session

from app.db.database import Base, SessionLocal, engine, get_db
from app.main import app


@pytest.fixture(scope="function")
def db_session() -> Session:
    """每个测试函数一个干净的内存数据库。"""
    Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
        Base.metadata.drop_all(bind=engine)


@pytest.fixture(scope="function")
def client(db_session: Session) -> TestClient:
    """FastAPI TestClient，覆盖 get_db 依赖用测试 session。"""

    def override_get_db():
        try:
            yield db_session
        finally:
            pass

    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()
```

注意：现在 `test_auth.py`、`test_subscription.py`、`test_key.py`、`test_logs.py` 里各自的 fixture 可以删除了（用 conftest 的）。但为了减少改动量，保留它们也行——pytest 同名 fixture 会优先用测试文件内的。**简化：让各测试文件删除自己的 fixture，统一用 conftest。**

每个测试文件删除 `db_session` 和 `client` fixture 定义（以及顶部的 `os.environ.setdefault` 行，因为 conftest 已经设了）。

- [ ] **Step 3: 写完整流程集成测试**

`server/tests/test_integration.py`：

```python
"""完整流程：注册 → 登录 → 查订阅 → 拉 Key → 上报日志。"""


def test_full_flow(client):
    # 1. 注册
    reg = client.post("/api/auth/register", json={
        "email": "flow@example.com",
        "password": "securepassword123",
    })
    assert reg.status_code == 201
    token = reg.json()["token"]

    headers = {"Authorization": f"Bearer {token}"}

    # 2. 查订阅
    sub = client.get("/api/subscription", headers=headers)
    assert sub.status_code == 200
    assert sub.json()["tier"] == "trial"
    assert sub.json()["status"] == "active"

    # 3. 拉 Key
    key = client.get("/api/key", headers=headers)
    assert key.status_code == 200
    assert key.json()["api_key"] == "sk-test-xxx"

    # 4. 上报日志
    logs = client.post("/api/logs", headers=headers, json={"logs": [
        {"timestamp": "2026-07-14T10:30:00Z", "level": "info", "event": "test"},
    ]})
    assert logs.status_code == 200
    assert logs.json()["accepted"] == 1

    # 5. 重新登录
    login = client.post("/api/auth/login", json={
        "email": "flow@example.com",
        "password": "securepassword123",
    })
    assert login.status_code == 200
    assert login.json()["token"]


def test_expired_subscription_blocks_key(client, db_session):
    """订阅过期后拉 Key 应返回 403。"""
    reg = client.post("/api/auth/register", json={
        "email": "expired@example.com",
        "password": "securepassword123",
    })
    token = reg.json()["token"]
    headers = {"Authorization": f"Bearer {token}"}

    # 把订阅设为已过期
    from datetime import datetime, timedelta, timezone
    from app.db.models import Subscription
    sub = db_session.query(Subscription).first()
    sub.expires_at = datetime.now(timezone.utc) - timedelta(days=1)
    db_session.commit()

    key = client.get("/api/key", headers=headers)
    assert key.status_code == 403


def test_health_endpoint(client):
    resp = client.get("/health")
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"
```

- [ ] **Step 4: 跑全部测试**

Run: `cd server && python -m pytest tests/ -v`
Expected: 全部 passed（约 25-30 个测试）

- [ ] **Step 5: Commit**

```bash
git add server/app/main.py server/tests/conftest.py server/tests/test_integration.py server/tests/test_auth.py server/tests/test_subscription.py server/tests/test_key.py server/tests/test_logs.py
git commit -m "feat(server): add CORS, integration tests, unify conftest fixtures"
```

---

### Task 9: Docker Compose 部署验证 + README

**Files:**
- Create: `server/README.md`
- Create: `server/.gitignore`
- Create: `server/alembic/versions/.gitkeep`

- [ ] **Step 1: 写 .gitignore**

`server/.gitignore`：

```
__pycache__/
*.pyc
.env
*.egg-info/
.pytest_cache/
alembic/versions/*.py
!alembic/versions/.gitkeep
```

- [ ] **Step 2: 创建 .gitkeep**

`server/alembic/versions/.gitkeep`：（空文件）

- [ ] **Step 3: 写 README.md**

`server/README.md`：

```markdown
# DeepExcel Server

DeepExcel SaaS 服务端 MVP。提供用户注册/登录、订阅校验、API Key 下发、客户端日志上报。

## 架构

- Agent loop 留客户端本地（Claude Agent SDK）
- 服务端只做"门卫 + 日志"
- Phase 1：直接下发真实 Claude API Key
- Phase 2（规划中）：短期 token + 服务端代理

## 快速启动

### 1. 准备环境变量

```bash
cp .env.example .env
# 编辑 .env，填入 CLAUDE_API_KEY 和 JWT_SECRET
```

### 2. Docker Compose 启动

```bash
docker-compose up -d --build
```

服务启动在 http://localhost:8000

### 3. 初始化数据库

首次启动后，进入容器跑 Alembic 迁移：

```bash
docker-compose exec api alembic upgrade head
```

### 4. 验证

```bash
curl http://localhost:8000/health
# {"status":"ok","version":"0.1.0"}
```

## API

| 方法 | 路径 | 说明 | 认证 |
|------|------|------|------|
| POST | /api/auth/register | 注册 | 无 |
| POST | /api/auth/login | 登录 | 无 |
| GET | /api/auth/me | 验证 JWT | JWT |
| GET | /api/subscription | 查订阅状态 | JWT |
| GET | /api/key | 拉 Claude API Key | JWT + 有效订阅 |
| POST | /api/logs | 批量上报日志 | JWT |
| GET | /health | 健康检查 | 无 |

## 本地开发（不用 Docker）

```bash
pip install -r requirements.txt

# 用 SQLite 跑测试
DATABASE_URL=sqlite:///./dev.db python -c "from app.db.database import Base; Base.metadata.create_all()"

uvicorn app.main:app --reload --port 8000
```

## 测试

```bash
python -m pytest tests/ -v
```

测试用 SQLite 内存数据库，无需 PostgreSQL。

## 技术栈

- FastAPI + Uvicorn
- SQLAlchemy 2.0 + Alembic
- PostgreSQL 16 + Redis 7
- bcrypt + python-jose (JWT)
- Docker Compose
```

- [ ] **Step 4: 验证 Docker Compose 构建（手动，需 Docker 环境）**

Run: `cd server && docker-compose build`
Expected: 构建成功，无错误

Run: `cd server && docker-compose up -d`
Expected: 三个容器（api、db、redis）都启动

Run: `cd server && docker-compose exec api alembic upgrade head`
Expected: 迁移成功，创建 users/subscriptions/logs 三张表

Run: `curl http://localhost:8000/health`
Expected: `{"status":"ok","version":"0.1.0"}`

Run: `curl -X POST http://localhost:8000/api/auth/register -H "Content-Type: application/json" -d '{"email":"test@example.com","password":"securepassword123"}'`
Expected: 201，返回 token

Run: `cd server && docker-compose down`
Expected: 容器停止

- [ ] **Step 5: Commit**

```bash
git add server/README.md server/.gitignore server/alembic/versions/.gitkeep
git commit -m "docs(server): add README, gitignore, deployment verification"
```

---

## Self-Review 检查

### Spec 覆盖检查

| Spec 要求 | 对应 Task |
|----------|-----------|
| POST /api/auth/register | Task 4 |
| POST /api/auth/login | Task 4 |
| GET /api/key | Task 6 |
| GET /api/subscription | Task 5 |
| POST /api/logs | Task 7 |
| User 模型 | Task 2 |
| Subscription 模型 | Task 2 |
| LogRecord 模型 | Task 2 |
| bcrypt cost=12 | Task 4 (service.py) |
| JWT 7 天有效 | Task 3 (config 默认值) |
| 30 天试用 | Task 4 (register_user 创建 Subscription) |
| Docker Compose | Task 1 + Task 9 |
| Alembic 迁移 | Task 2 + Task 9 |
| 测试（单元 + 集成 + 安全） | Task 3-8 |
| CORS | Task 8 |
| .env.example | Task 1 |
| 不做支付/管理后台/邮件验证 | ✓ 全部未涉及 |

### Placeholder 扫描

- 无 "TBD"、"TODO"、"implement later"
- 所有代码块完整，无省略
- 所有命令都有 expected output

### 类型一致性检查

- `create_access_token(user_id: str, tier: str) -> tuple[str, datetime]` — Task 3 定义，Task 4 调用 ✅
- `verify_token(token: str) -> dict` — Task 3 定义，Task 4 (dependencies.py) 调用 ✅
- `get_current_user_id` 依赖 — Task 4 定义，Task 5/6/7 调用 ✅
- `is_subscription_active(db, user_id) -> bool` — Task 5 定义，Task 6 调用 ✅
- `get_db()` 依赖 — Task 2 定义，Task 4-7 调用 ✅
- `Settings` 类字段 — Task 1 定义，Task 2/3/4/6 读取 ✅

---

## Execution Handoff

**Plan complete and saved to `docs/superpowers/plans/2026-07-14-saas-server-mvp.md`.**

用户明确说"我后面叫你开发时再开发，近期估计还不开发"，所以**不立即执行**。当用户准备好开发时，两个执行选项：

**1. Subagent-Driven（推荐）** - 每个 Task 派一个 fresh subagent，Task 间 review，快速迭代

**2. Inline Execution** - 在当前 session 用 executing-plans skill 批量执行，checkpoint review

等待用户指令再开始。
