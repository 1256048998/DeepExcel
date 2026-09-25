"""The update feed.

What matters here is small: serve the signed bytes untouched, refuse to serve
anything that is obviously not a manifest, and never let a channel name reach
the filesystem. The server intentionally does not verify the signature -- the
client does, which is what makes a compromised server unable to push code -- so
these tests assert relaying, not trust.
"""

from __future__ import annotations

import json

import pytest

from app.routers import updates as updates_router

MANIFEST = {
    "schema": 1,
    "key_id": "fcf2af18a20c536a",
    "signature": "c2lnbmF0dXJl",
    "payload": "cGF5bG9hZA==",
}


@pytest.fixture
def feed(make_client, tmp_path):
    """A client whose update feed is a directory we control."""
    directory = tmp_path / "updates"
    directory.mkdir()

    def _publish(channel: str, manifest) -> None:
        body = manifest if isinstance(manifest, str) else json.dumps(manifest)
        (directory / f"{channel}.json").write_text(body, encoding="utf-8")
        updates_router.reset_cache()

    client = make_client(UPDATE_MANIFEST_DIR=str(directory))
    return client, _publish, directory


def test_serves_the_signed_manifest_byte_for_byte(feed):
    client, publish, directory = feed
    publish("stable", MANIFEST)
    on_disk = (directory / "stable.json").read_text(encoding="utf-8")

    response = client.get("/api/v1/updates/latest")

    assert response.status_code == 200, response.text
    # Byte-for-byte: re-encoding the JSON would change the bytes under the
    # signature and every client would reject the result.
    assert response.text == on_disk
    assert response.json() == MANIFEST


def test_requires_no_authentication(feed):
    client, publish, _ = feed
    publish("stable", MANIFEST)

    # A user whose session expired still has to be able to receive a fix.
    assert client.get("/api/v1/updates/latest").status_code == 200


def test_channels_are_served_separately(feed):
    client, publish, _ = feed
    publish("stable", MANIFEST)
    publish("beta", dict(MANIFEST, key_id="beta0000beta0000"))

    assert client.get("/api/v1/updates/latest?channel=beta").json()["key_id"] == "beta0000beta0000"
    assert client.get("/api/v1/updates/latest").json()["key_id"] == MANIFEST["key_id"]


def test_unpublished_channel_is_a_404_not_an_error(feed):
    client, publish, _ = feed
    publish("stable", MANIFEST)

    response = client.get("/api/v1/updates/latest?channel=beta")

    assert response.status_code == 404
    assert response.json()["detail"]["reason"] == "no_release"


@pytest.mark.parametrize(
    "channel",
    ["../secrets", "..", "stable/../../etc/passwd", "a" * 33, "", "Stable!", "sta ble"],
)
def test_a_channel_name_can_never_reach_the_filesystem(feed, channel):
    client, publish, _ = feed
    publish("stable", MANIFEST)

    response = client.get("/api/v1/updates/latest", params={"channel": channel})

    assert response.status_code == 400, response.text
    assert response.json()["detail"]["reason"] == "bad_channel"


def test_an_unconfigured_deployment_says_so(make_client):
    client = make_client()

    response = client.get("/api/v1/updates/latest")

    assert response.status_code == 503
    assert response.json()["detail"]["reason"] == "updates_not_configured"


@pytest.mark.parametrize(
    "body",
    [
        "not json at all",
        "[1, 2, 3]",
        json.dumps({"schema": 1, "key_id": "x"}),          # no signature or payload
        json.dumps({"signature": "a", "payload": "b"}),     # no schema or key_id
    ],
)
def test_refuses_to_serve_a_file_that_is_not_a_manifest(feed, body):
    """Catches the operator who copied the wrong file.

    Serving junk would not endanger anyone -- the client rejects it -- but it
    would look like a working deployment while every client silently stopped
    updating.
    """
    client, publish, _ = feed
    publish("stable", body)

    response = client.get("/api/v1/updates/latest")

    assert response.status_code == 503
    assert response.json()["detail"]["reason"] == "manifest_invalid"


def test_refuses_an_implausibly_large_manifest(feed):
    client, publish, directory = feed
    publish("stable", MANIFEST)
    (directory / "stable.json").write_text(
        "x" * (updates_router.MAX_MANIFEST_BYTES + 1), encoding="utf-8")
    updates_router.reset_cache()

    assert client.get("/api/v1/updates/latest").status_code == 503


def test_a_new_release_is_picked_up_without_a_restart(feed):
    client, publish, directory = feed
    publish("stable", MANIFEST)
    assert client.get("/api/v1/updates/latest").json()["key_id"] == MANIFEST["key_id"]

    # Publishing is a file drop; an operator should not have to restart the
    # service to ship a fix.
    publish("stable", dict(MANIFEST, key_id="1111222233334444"))

    assert client.get("/api/v1/updates/latest").json()["key_id"] == "1111222233334444"


def test_meta_reports_whether_a_feed_exists(feed, make_client):
    client, publish, _ = feed
    publish("stable", MANIFEST)

    assert client.get("/api/v1/meta").json()["update_feed_available"] is True
    assert make_client().get("/api/v1/meta").json()["update_feed_available"] is False


def test_response_is_cacheable(feed):
    client, publish, _ = feed
    publish("stable", MANIFEST)

    response = client.get("/api/v1/updates/latest")

    # The feed is polled by every client on every start; without this it is a
    # request per client per session straight to the origin.
    assert "max-age" in response.headers["cache-control"]


def test_knowledge_pack_is_relayed_byte_for_byte(feed):
    client, _, directory = feed
    body = json.dumps(dict(MANIFEST, payload="a" * 200_000))
    (directory / updates_router.KNOWLEDGE_PACK_FILE).write_text(body, encoding="utf-8")
    updates_router.reset_cache()

    response = client.get("/api/v1/updates/knowledge")

    assert response.status_code == 200, response.text
    # Larger than a manifest may be: the knowledge limit is its own.
    assert response.text == body


def test_knowledge_pack_unpublished_is_a_404(feed):
    client, _, _ = feed
    response = client.get("/api/v1/updates/knowledge")
    assert response.status_code == 404
    assert response.json()["detail"]["reason"] == "no_release"


def test_knowledge_pack_is_never_served_as_a_channel(feed):
    client, _, directory = feed
    (directory / updates_router.KNOWLEDGE_PACK_FILE).write_text(json.dumps(MANIFEST), encoding="utf-8")
    updates_router.reset_cache()

    response = client.get("/api/v1/updates/latest?channel=knowledge_pack")

    assert response.status_code == 400


def test_knowledge_pack_that_is_not_a_signed_envelope_is_refused(feed):
    client, _, directory = feed
    (directory / updates_router.KNOWLEDGE_PACK_FILE).write_text('{"skills": []}', encoding="utf-8")
    updates_router.reset_cache()

    response = client.get("/api/v1/updates/knowledge")

    assert response.status_code == 503
    assert response.json()["detail"]["reason"] == "manifest_invalid"


def test_knowledge_endpoint_without_feed_dir_is_unconfigured(make_client):
    response = make_client().get("/api/v1/updates/knowledge")
    assert response.status_code == 503
    assert response.json()["detail"]["reason"] == "updates_not_configured"
