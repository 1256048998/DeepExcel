"""Skill sync and sharing.

The property that matters most is scrubbing. A skill's parameter defaults are
values captured from a real run, so they carry real local paths and real
workbook names. Uploading them verbatim would put someone's file layout in a
database; sharing would hand it to a stranger.
"""

from __future__ import annotations

from app.routers.skills import scrub, scrub_skill
from tests.conftest import auth_headers, register


def _skill(**overrides) -> dict:
    document = {
        "id": "abc123",
        "name": "月度销售报表",
        "description": None,
        "original_request": "把 A1:F200 按销售额降序排列",
        "parameters": [
            {"name": "range", "kind": "Range", "label": "数据区域", "default_value": "A1:F200"},
        ],
        "steps": [
            {"tool": "sort_data", "arguments": {"address": "{{range}}"}},
            {"tool": "create_chart", "arguments": {"address": "{{range}}"}},
        ],
        "run_count": 3,
    }
    document.update(overrides)
    return document


# ---------------------------------------------------------------------------
# Scrubbing
# ---------------------------------------------------------------------------

def test_windows_paths_are_removed():
    assert "alice" not in scrub(r"读取 C:\Users\alice\2026财务预算.xlsx 的数据")
    assert "alice" not in scrub(r"\\fileserver\share\alice\budget.xlsx")


def test_workbook_names_are_removed():
    # Even without a path, a file name can identify a project, a client or a
    # quarter's results.
    cleaned = scrub("汇总 2026年Q3并购尽调.xlsx 的数据")
    assert "并购尽调" not in cleaned
    assert "xlsx" not in cleaned


def test_structure_is_preserved():
    # Tool names and placeholders are structure, not data -- scrubbing them
    # would destroy the skill.
    cleaned = scrub("把 {{range}} 按销售额降序排列")
    assert cleaned == "把 {{range}} 按销售额降序排列"


def test_file_parameter_defaults_never_survive_upload():
    document = _skill(
        parameters=[
            {"name": "file", "kind": "File", "label": "文件",
             "default_value": r"C:\Users\alice\2026财务预算.xlsx"},
            {"name": "range", "kind": "Range", "label": "数据区域", "default_value": "A1:F200"},
        ]
    )
    scrubbed = scrub_skill(document)

    # The recipient's file lives somewhere else anyway, so the default is
    # worthless to them and harmful to keep.
    assert scrubbed["parameters"][0]["default_value"] == ""
    # A range is not identifying and is genuinely useful as a default.
    assert scrubbed["parameters"][1]["default_value"] == "A1:F200"


def test_scrubbing_reaches_the_original_request_and_step_arguments():
    document = _skill(
        original_request=r"读取 C:\Users\bob\salary.xlsx 并汇总",
        steps=[{"tool": "read_range", "arguments": {"path": r"D:\HR\salary.xlsx"}}],
    )
    scrubbed = scrub_skill(document)

    assert "bob" not in scrubbed["original_request"]
    assert "salary" not in scrubbed["original_request"]
    assert "salary" not in scrubbed["steps"][0]["arguments"]["path"]


def test_scrubbing_tolerates_malformed_documents():
    # A hand-edited or hostile document must not crash the endpoint.
    scrub_skill({"parameters": ["not a dict"], "steps": [None, {"arguments": "not a dict"}]})
    scrub_skill({})


# ---------------------------------------------------------------------------
# Sync
# ---------------------------------------------------------------------------

def test_upload_and_fetch_round_trip(client):
    tokens = register(client)
    headers = auth_headers(tokens)

    response = client.put(
        "/api/v1/skills",
        json={"skill_id": "abc123", "name": "月度销售报表", "document": _skill()},
        headers=headers,
    )
    assert response.status_code == 200
    assert response.json()["share_code"] is None  # 同步不等于公开

    fetched = client.get("/api/v1/skills/abc123", headers=headers).json()
    assert fetched["name"] == "月度销售报表"
    assert len(fetched["steps"]) == 2


def test_reupload_updates_rather_than_duplicates(client):
    tokens = register(client)
    headers = auth_headers(tokens)

    client.put("/api/v1/skills",
               json={"skill_id": "abc123", "name": "旧名", "document": _skill()}, headers=headers)
    client.put("/api/v1/skills",
               json={"skill_id": "abc123", "name": "新名", "document": _skill(name="新名")},
               headers=headers)

    listed = client.get("/api/v1/skills", headers=headers).json()
    assert len(listed) == 1
    assert listed[0]["name"] == "新名"


def test_uploaded_paths_are_scrubbed_server_side(client):
    # The client scrubs too, but it is unsigned and modifiable, so its filtering
    # is a convenience rather than a guarantee.
    tokens = register(client)
    headers = auth_headers(tokens)

    client.put(
        "/api/v1/skills",
        json={
            "skill_id": "leaky",
            "name": r"处理 C:\Users\alice\预算.xlsx",
            "document": _skill(original_request=r"读取 C:\Users\alice\预算.xlsx"),
        },
        headers=headers,
    )

    stored = client.get("/api/v1/skills/leaky", headers=headers).json()
    assert "alice" not in stored["original_request"]
    assert "alice" not in client.get("/api/v1/skills", headers=headers).json()[0]["name"]


def test_skills_are_private_to_their_owner(client):
    first = register(client, email="first@deepexcel-qa.com")
    second = register(client, email="second@deepexcel-qa.com")

    client.put("/api/v1/skills",
               json={"skill_id": "mine", "name": "我的", "document": _skill()},
               headers=auth_headers(first))

    assert client.get("/api/v1/skills", headers=auth_headers(second)).json() == []
    assert client.get("/api/v1/skills/mine", headers=auth_headers(second)).status_code == 404


def test_deleting_removes_it(client):
    tokens = register(client)
    headers = auth_headers(tokens)
    client.put("/api/v1/skills",
               json={"skill_id": "temp", "name": "临时", "document": _skill()}, headers=headers)

    assert client.delete("/api/v1/skills/temp", headers=headers).status_code == 204
    assert client.get("/api/v1/skills", headers=headers).json() == []


def test_oversized_skill_is_rejected(client):
    tokens = register(client)
    huge = _skill(steps=[{"tool": "x", "arguments": {"v": "y" * 200}} for _ in range(1000)])

    response = client.put(
        "/api/v1/skills",
        json={"skill_id": "huge", "name": "巨大", "document": huge},
        headers=auth_headers(tokens),
    )
    assert response.status_code == 413


# ---------------------------------------------------------------------------
# Sharing
# ---------------------------------------------------------------------------

def test_sharing_is_opt_in_and_reversible(client):
    owner = register(client, email="owner@deepexcel-qa.com")
    other = register(client, email="other@deepexcel-qa.com")

    client.put("/api/v1/skills",
               json={"skill_id": "s1", "name": "报表", "document": _skill()},
               headers=auth_headers(owner))

    code = client.post("/api/v1/skills/s1/share", headers=auth_headers(owner)).json()["share_code"]
    imported = client.get(f"/api/v1/skills/shared/{code}", headers=auth_headers(other))
    assert imported.status_code == 200
    # 分享返回的是技能文档本身，它自带名字
    assert imported.json()["name"] == "月度销售报表"
    assert len(imported.json()["steps"]) == 2

    client.post("/api/v1/skills/s1/unshare", headers=auth_headers(owner))
    assert client.get(f"/api/v1/skills/shared/{code}", headers=auth_headers(other)).status_code == 404


def test_importing_assigns_a_fresh_id(client):
    owner = register(client, email="owner@deepexcel-qa.com")
    other = register(client, email="other@deepexcel-qa.com")

    client.put("/api/v1/skills",
               json={"skill_id": "s1", "name": "报表", "document": _skill()},
               headers=auth_headers(owner))
    code = client.post("/api/v1/skills/s1/share", headers=auth_headers(owner)).json()["share_code"]

    imported = client.get(f"/api/v1/skills/shared/{code}", headers=auth_headers(other)).json()
    # Otherwise importing would overwrite a skill the recipient already has with
    # the same id.
    assert imported["id"] != "abc123"
    assert len(imported["id"]) == 32


def test_sharing_the_same_skill_twice_keeps_one_code(client):
    tokens = register(client)
    headers = auth_headers(tokens)
    client.put("/api/v1/skills",
               json={"skill_id": "s1", "name": "报表", "document": _skill()}, headers=headers)

    first = client.post("/api/v1/skills/s1/share", headers=headers).json()["share_code"]
    second = client.post("/api/v1/skills/s1/share", headers=headers).json()["share_code"]
    # A new code on every click would invalidate links already sent.
    assert first == second


def test_share_import_requires_an_account(client):
    # Codes are short; an open endpoint would make them enumerable.
    assert client.get("/api/v1/skills/shared/anything").status_code == 401


def test_unknown_share_code_is_not_found(client):
    tokens = register(client)
    assert client.get("/api/v1/skills/shared/nope", headers=auth_headers(tokens)).status_code == 404
