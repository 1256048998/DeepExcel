"""Plans, orders and activation.

The payment channel is not integrated. WeChat Pay and Alipay both require a
verified merchant account, which takes weeks to obtain, so the money side is
deliberately left as one well-defined seam: an order is created here, and
something marks it paid. Today that something is an operator; later it is a
payment callback. Everything downstream of "order is paid" -- entitlement
upgrade, expiry, quota reset -- works now and is tested now, so turning on a
channel does not also mean debugging entitlements for the first time.
"""

from __future__ import annotations

import datetime as dt
import secrets

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from ..db import get_db
from ..deps import get_current_admin, get_current_user
from ..models import (
    Admin,
    AuditLog,
    Entitlement,
    EntitlementStatus,
    Order,
    OrderStatus,
    Plan,
    RoutingMode,
    User,
    utcnow,
)
from ..schemas import OrderCreate, OrderView, PlanView

router = APIRouter(tags=["billing"])


# Prices in cents. Per-task rather than per-token because a user can predict how
# many tables they will process in a month and cannot predict tokens.
CATALOG: dict[str, dict] = {
    "free": {
        "label": "免费",
        "price_cents": 0,
        "task_limit": 50,
        "routing": RoutingMode.HOSTED,
        "summary": "每月 50 次任务，用于试用",
    },
    "pro": {
        "label": "专业版",
        "price_cents": 3900,
        "task_limit": 1000,
        "routing": RoutingMode.HOSTED,
        "summary": "每月 1000 次任务",
    },
    "team": {
        "label": "团队版",
        "price_cents": 9900,
        "task_limit": None,
        "routing": RoutingMode.HOSTED,
        "summary": "不限次数，技能共享与集中管理",
    },
    "byok": {
        "label": "自带密钥",
        "price_cents": 1900,
        "task_limit": None,
        # Traffic never touches our infrastructure, so there is nothing to
        # meter. This tier buys the software and the skill library, not tokens.
        "routing": RoutingMode.BYOK,
        "summary": "使用自己的 API Key，模型流量不经过 DeepExcel",
    },
}


@router.get("/api/v1/plans", response_model=list[PlanView])
def list_plans() -> list[PlanView]:
    """Published so the client can render pricing without shipping a build."""
    return [
        PlanView(
            plan=key,
            label=value["label"],
            price_cents=value["price_cents"],
            currency="CNY",
            task_limit=value["task_limit"],
            summary=value["summary"],
        )
        for key, value in CATALOG.items()
    ]


@router.post("/api/v1/orders", response_model=OrderView, status_code=status.HTTP_201_CREATED)
def create_order(
    payload: OrderCreate,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> Order:
    plan = CATALOG.get(payload.plan)
    if plan is None or plan["price_cents"] == 0:
        # The free tier is granted, not purchased.
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="无效的套餐")

    order = Order(
        order_no=utcnow().strftime("%Y%m%d") + secrets.token_hex(6).upper(),
        user_id=user.id,
        plan=Plan(payload.plan),
        months=payload.months,
        amount_cents=plan["price_cents"] * payload.months,
        currency="CNY",
        status=OrderStatus.PENDING,
    )
    db.add(order)
    db.flush()
    return order


@router.get("/api/v1/orders", response_model=list[OrderView])
def my_orders(
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> list[Order]:
    return list(
        db.scalars(
            select(Order).where(Order.user_id == user.id).order_by(Order.created_at.desc()).limit(50)
        )
    )


def activate(db: Session, order: Order) -> Entitlement:
    """Applies a paid order to the user's entitlement.

    The seam a payment callback will call. Extending rather than replacing the
    expiry means paying twice before the first period ends adds time instead of
    silently discarding what was already bought.
    """
    entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == order.user_id))
    if entitlement is None:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="用户没有权益记录")

    plan = CATALOG[order.plan.value]
    now = utcnow()
    current_expiry = entitlement.expires_at
    if current_expiry is not None and current_expiry.tzinfo is None:
        current_expiry = current_expiry.replace(tzinfo=dt.timezone.utc)

    base = current_expiry if current_expiry and current_expiry > now else now
    entitlement.plan = order.plan
    entitlement.status = EntitlementStatus.ACTIVE
    entitlement.routing_mode = plan["routing"]
    entitlement.task_limit = plan["task_limit"]
    entitlement.expires_at = base + dt.timedelta(days=30 * order.months)
    # A new paid period starts with a clean quota.
    entitlement.tasks_used = 0
    entitlement.period_started_at = now

    order.status = OrderStatus.PAID
    order.paid_at = now
    return entitlement


@router.post("/admin/api/orders/{order_no}/mark-paid", response_model=OrderView)
def mark_paid(
    order_no: str,
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> Order:
    """Manual activation, until a payment channel exists.

    Every use is audited: an operator being able to grant a paid plan is exactly
    the kind of power that needs a record of who used it.
    """
    order = db.scalar(select(Order).where(Order.order_no == order_no))
    if order is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="订单不存在")
    if order.status is OrderStatus.PAID:
        return order

    order.channel = "manual"
    activate(db, order)
    db.add(
        AuditLog(
            admin_id=admin.id,
            action="order.mark_paid",
            target=order.order_no,
            detail=f'{{"plan":"{order.plan.value}","months":{order.months}}}',
        )
    )
    return order


@router.get("/admin/api/orders", response_model=list[OrderView])
def list_orders(
    admin: Admin = Depends(get_current_admin),
    db: Session = Depends(get_db),
) -> list[Order]:
    return list(db.scalars(select(Order).order_by(Order.created_at.desc()).limit(200)))
