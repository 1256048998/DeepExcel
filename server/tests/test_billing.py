"""Plans, orders and entitlement activation.

The payment channel is not integrated -- a merchant account takes weeks. What is
tested here is everything downstream of "an order became paid", so that turning
on a channel later does not also mean debugging entitlements for the first time.
"""

from __future__ import annotations

import datetime as dt

from sqlalchemy import select

from app.db import get_session_factory
from app.models import Entitlement, RoutingMode
from tests.conftest import admin_token, auth_headers, register


def test_plans_are_published_so_pricing_needs_no_client_release(client):
    plans = {plan["plan"]: plan for plan in client.get("/api/v1/plans").json()}

    assert plans["free"]["price_cents"] == 0
    assert plans["pro"]["task_limit"] == 1000
    assert plans["team"]["task_limit"] is None
    # BYOK buys the software and the skill library, not tokens.
    assert plans["byok"]["price_cents"] < plans["pro"]["price_cents"]


def test_creating_an_order_does_not_grant_anything(client):
    tokens = register(client)
    order = client.post(
        "/api/v1/orders", json={"plan": "pro", "months": 1}, headers=auth_headers(tokens)
    )

    assert order.status_code == 201
    body = order.json()
    assert body["status"] == "pending"
    assert body["amount_cents"] == 3900

    # Until it is paid the user is still on the beta plan.
    me = client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()
    assert me["entitlement"]["plan"] == "beta"


def test_multi_month_orders_are_priced_linearly(client):
    tokens = register(client)
    order = client.post(
        "/api/v1/orders", json={"plan": "pro", "months": 12}, headers=auth_headers(tokens)
    ).json()
    assert order["amount_cents"] == 3900 * 12


def test_free_plan_cannot_be_ordered(client):
    tokens = register(client)
    # It is granted, not purchased; accepting an order for it would create a
    # zero-value row that looks like a real purchase.
    response = client.post(
        "/api/v1/orders", json={"plan": "free"}, headers=auth_headers(tokens)
    )
    assert response.status_code == 422


def test_marking_paid_upgrades_the_entitlement(admin_client):
    tokens = register(admin_client)
    user_id = admin_client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    order = admin_client.post(
        "/api/v1/orders", json={"plan": "pro", "months": 1}, headers=auth_headers(tokens)
    ).json()

    paid = admin_client.post(
        f"/admin/api/orders/{order['order_no']}/mark-paid", headers=admin_token(admin_client)
    )
    assert paid.status_code == 200
    assert paid.json()["status"] == "paid"

    with get_session_factory()() as db:
        entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))
        assert entitlement.plan.value == "pro"
        assert entitlement.task_limit == 1000
        assert entitlement.routing_mode is RoutingMode.HOSTED
        assert entitlement.expires_at is not None
        # A new paid period starts with a clean quota.
        assert entitlement.tasks_used == 0


def test_byok_plan_does_not_route_through_the_proxy(admin_client):
    tokens = register(admin_client)
    user_id = admin_client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    order = admin_client.post(
        "/api/v1/orders", json={"plan": "byok"}, headers=auth_headers(tokens)
    ).json()
    admin_client.post(
        f"/admin/api/orders/{order['order_no']}/mark-paid", headers=admin_token(admin_client)
    )

    with get_session_factory()() as db:
        entitlement = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id))
        # Their traffic never touches us, so there is nothing to meter.
        assert entitlement.routing_mode is RoutingMode.BYOK
        assert entitlement.task_limit is None


def test_renewing_early_extends_rather_than_discards(admin_client):
    tokens = register(admin_client)
    user_id = admin_client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    ops = admin_token(admin_client)

    first = admin_client.post(
        "/api/v1/orders", json={"plan": "pro", "months": 1}, headers=auth_headers(tokens)
    ).json()
    admin_client.post(f"/admin/api/orders/{first['order_no']}/mark-paid", headers=ops)
    with get_session_factory()() as db:
        after_first = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id)).expires_at

    second = admin_client.post(
        "/api/v1/orders", json={"plan": "pro", "months": 1}, headers=auth_headers(tokens)
    ).json()
    admin_client.post(f"/admin/api/orders/{second['order_no']}/mark-paid", headers=ops)
    with get_session_factory()() as db:
        after_second = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id)).expires_at

    if after_first.tzinfo is None:
        after_first = after_first.replace(tzinfo=dt.timezone.utc)
        after_second = after_second.replace(tzinfo=dt.timezone.utc)

    # Paying twice before the first period ends must add time, not throw away
    # what was already bought.
    assert after_second > after_first
    assert (after_second - after_first).days >= 29


def test_marking_paid_twice_is_idempotent(admin_client):
    tokens = register(admin_client)
    user_id = admin_client.get("/api/v1/auth/me", headers=auth_headers(tokens)).json()["id"]
    ops = admin_token(admin_client)
    order = admin_client.post(
        "/api/v1/orders", json={"plan": "pro"}, headers=auth_headers(tokens)
    ).json()

    admin_client.post(f"/admin/api/orders/{order['order_no']}/mark-paid", headers=ops)
    with get_session_factory()() as db:
        first_expiry = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id)).expires_at

    # A payment callback can fire twice; that must not grant a second period.
    admin_client.post(f"/admin/api/orders/{order['order_no']}/mark-paid", headers=ops)
    with get_session_factory()() as db:
        second_expiry = db.scalar(select(Entitlement).where(Entitlement.user_id == user_id)).expires_at

    assert first_expiry == second_expiry


def test_activation_is_audited(admin_client):
    tokens = register(admin_client)
    ops = admin_token(admin_client)
    order = admin_client.post(
        "/api/v1/orders", json={"plan": "pro"}, headers=auth_headers(tokens)
    ).json()
    admin_client.post(f"/admin/api/orders/{order['order_no']}/mark-paid", headers=ops)

    actions = [row["action"] for row in admin_client.get("/admin/api/audit", headers=ops).json()]
    # Granting a paid plan by hand is exactly the power that needs a record of
    # who used it.
    assert "order.mark_paid" in actions


def test_orders_are_scoped_to_their_owner(client):
    first = register(client, email="first@deepexcel-qa.com")
    second = register(client, email="second@deepexcel-qa.com")

    client.post("/api/v1/orders", json={"plan": "pro"}, headers=auth_headers(first))

    assert len(client.get("/api/v1/orders", headers=auth_headers(first)).json()) == 1
    assert client.get("/api/v1/orders", headers=auth_headers(second)).json() == []


def test_unknown_order_cannot_be_activated(admin_client):
    response = admin_client.post(
        "/admin/api/orders/NOPE/mark-paid", headers=admin_token(admin_client)
    )
    assert response.status_code == 404


def test_users_cannot_activate_their_own_orders(client):
    tokens = register(client)
    order = client.post(
        "/api/v1/orders", json={"plan": "team"}, headers=auth_headers(tokens)
    ).json()

    response = client.post(
        f"/admin/api/orders/{order['order_no']}/mark-paid", headers=auth_headers(tokens)
    )
    assert response.status_code == 401
