#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sign DeepExcel update manifests, and keep the compiled-in public key honest.

The auto-updater's whole security model is this file plus
``src/DeepExcel.AddIn/Updates/UpdateSigning.cs``:

  * the private key signs a release manifest here, offline;
  * the public key is compiled into the client;
  * the server only relays an already-signed manifest, and never holds the
    private key -- so taking over the server does not let anyone push an update.

Commands
--------
    python scripts/update_signing.py genkey --out D:/offline/deepexcel-update.pem
    python scripts/update_signing.py pubkey --key <pem>
    python scripts/update_signing.py sign --key <pem> --payload payload.json --out update.json
    python scripts/update_signing.py verify --manifest update.json

``genkey`` rewrites the two constants in UpdateSigning.cs in place, so the
compiled-in public key cannot drift from the private key that signs releases.

Interop note, the one that actually bites: .NET fixes the RSA-PSS salt length
at the digest size. ``padding.PSS.MAX_LENGTH`` produces signatures that .NET
rejects without explanation, so DIGEST_LENGTH is not a preference here.
"""

import argparse
import base64
import hashlib
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SIGNING_CS = os.path.join(ROOT, "src", "DeepExcel.AddIn", "Updates", "UpdateSigning.cs")

SCHEMA = 1
PRODUCT = "DeepExcel"
KEY_BITS = 4096
PUBLIC_EXPONENT = 65537

_BEGIN = "// BEGIN GENERATED UPDATE KEY"
_END = "// END GENERATED UPDATE KEY"


class SigningError(RuntimeError):
    pass


def _require_cryptography():
    try:
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import padding, rsa
    except ImportError as exc:  # pragma: no cover - environment problem
        raise SigningError(
            "The 'cryptography' package is required for update signing:\n"
            "    python -m pip install cryptography"
        ) from exc
    return hashes, serialization, padding, rsa


# ---------------------------------------------------------------- key material


def public_key_bytes(public_numbers):
    """The exact byte encodings the C# side imports into RSAParameters.

    Modulus is fixed-width big-endian (no leading zero stripped), exponent is
    minimal big-endian. Both sides derive the key id from these bytes, so any
    disagreement here shows up as a key-id mismatch rather than as a signature
    that mysteriously fails.
    """
    modulus = public_numbers.n.to_bytes((public_numbers.n.bit_length() + 7) // 8, "big")
    exponent = public_numbers.e.to_bytes((public_numbers.e.bit_length() + 7) // 8, "big")
    return modulus, exponent


def compute_key_id(modulus, exponent):
    return hashlib.sha256(modulus + exponent).hexdigest()[:16]


def load_private_key(path):
    hashes, serialization, padding, rsa = _require_cryptography()
    if not os.path.isfile(path):
        raise SigningError("Private key not found: %s" % path)
    with open(path, "rb") as stream:
        key = serialization.load_pem_private_key(stream.read(), password=None)
    if not isinstance(key, rsa.RSAPrivateKey):
        raise SigningError("Not an RSA private key: %s" % path)
    if key.key_size < 2048:
        raise SigningError("Refusing a %d-bit signing key; 2048 is the floor" % key.key_size)
    return key


def describe_key(path):
    key = load_private_key(path)
    modulus, exponent = public_key_bytes(key.public_key().public_numbers())
    return {
        "bits": key.key_size,
        "key_id": compute_key_id(modulus, exponent),
        "modulus_b64": base64.b64encode(modulus).decode("ascii"),
        "exponent_b64": base64.b64encode(exponent).decode("ascii"),
    }


# ------------------------------------------------------- compiled-in constants


def read_embedded_key(source_path=SIGNING_CS):
    """Returns the public key currently compiled into the client, or None."""
    with open(source_path, "r", encoding="utf-8") as stream:
        text = stream.read()
    modulus = re.search(r'ModulusBase64\s*=\s*"([^"]*)"', text)
    exponent = re.search(r'ExponentBase64\s*=\s*"([^"]*)"', text)
    if not modulus or not exponent or not modulus.group(1) or not exponent.group(1):
        return None
    raw_modulus = base64.b64decode(modulus.group(1))
    raw_exponent = base64.b64decode(exponent.group(1))
    return {
        "key_id": compute_key_id(raw_modulus, raw_exponent),
        "modulus_b64": modulus.group(1),
        "exponent_b64": exponent.group(1),
    }


def write_embedded_key(modulus_b64, exponent_b64, source_path=SIGNING_CS):
    """Rewrites the generated block in UpdateSigning.cs."""
    with open(source_path, "r", encoding="utf-8") as stream:
        text = stream.read()
    if _BEGIN not in text or _END not in text:
        raise SigningError("Generated key markers are missing from %s" % source_path)

    indent = " " * 8
    replacement = (
        "%s%s -- scripts/update_signing.py rewrites these two lines\n"
        '%spublic const string ModulusBase64 = "%s";\n'
        '%spublic const string ExponentBase64 = "%s";\n'
        "%s%s"
        % (indent, _BEGIN,
           indent, modulus_b64,
           indent, exponent_b64,
           indent, _END)
    )
    pattern = re.compile(
        re.escape(indent) + re.escape(_BEGIN) + r".*?" + re.escape(_END),
        re.DOTALL,
    )
    updated, count = pattern.subn(lambda _m: replacement, text, count=1)
    if count != 1:
        raise SigningError("Could not locate the generated key block in %s" % source_path)
    with open(source_path, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(updated)


# ------------------------------------------------------------------- manifests


def build_payload(version, url, sha256, size, channel="stable",
                  released_at=None, notes="", minimum_upgradable_version=None):
    if len(sha256 or "") != 64 or any(c not in "0123456789abcdef" for c in sha256.lower()):
        raise SigningError("sha256 must be 64 hex characters")
    if not (url or "").startswith("https://"):
        raise SigningError("Update URL must be https")
    if size <= 0:
        raise SigningError("size must be positive")

    payload = {
        "schema": SCHEMA,
        "product": PRODUCT,
        "channel": channel,
        "version": version,
        "url": url,
        "sha256": sha256.lower(),
        "size": int(size),
        "released_at": released_at or "",
        "notes": notes or "",
    }
    if minimum_upgradable_version:
        payload["minimum_upgradable_version"] = minimum_upgradable_version
    return payload


def sign_payload(private_key_path, payload):
    """Wraps a payload into a signed manifest.

    The payload is embedded as opaque base64 of the exact bytes that were
    signed. The client verifies over those bytes and only then parses them, so
    neither side ever has to agree on a canonical JSON form.
    """
    hashes, serialization, padding, rsa = _require_cryptography()
    key = load_private_key(private_key_path)
    modulus, exponent = public_key_bytes(key.public_key().public_numbers())

    payload_bytes = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")

    signature = key.sign(
        payload_bytes,
        padding.PSS(mgf=padding.MGF1(hashes.SHA256()),
                    salt_length=hashes.SHA256.digest_size),
        hashes.SHA256(),
    )

    return {
        "schema": SCHEMA,
        "key_id": compute_key_id(modulus, exponent),
        "signature": base64.b64encode(signature).decode("ascii"),
        "payload": base64.b64encode(payload_bytes).decode("ascii"),
    }


def verify_manifest(manifest, modulus_b64, exponent_b64):
    """Mirrors the client check, so a release can be proved good before shipping."""
    hashes, serialization, padding, rsa = _require_cryptography()

    if manifest.get("schema") != SCHEMA:
        raise SigningError("Unsupported manifest schema: %r" % manifest.get("schema"))

    modulus = base64.b64decode(modulus_b64)
    exponent = base64.b64decode(exponent_b64)
    expected_key_id = compute_key_id(modulus, exponent)
    if manifest.get("key_id") != expected_key_id:
        raise SigningError(
            "Manifest key_id %r does not match the public key %r"
            % (manifest.get("key_id"), expected_key_id)
        )

    payload_bytes = base64.b64decode(manifest["payload"])
    signature = base64.b64decode(manifest["signature"])

    public_key = rsa.RSAPublicNumbers(
        int.from_bytes(exponent, "big"), int.from_bytes(modulus, "big")
    ).public_key()
    from cryptography.exceptions import InvalidSignature

    try:
        public_key.verify(
            signature,
            payload_bytes,
            padding.PSS(mgf=padding.MGF1(hashes.SHA256()),
                        salt_length=hashes.SHA256.digest_size),
            hashes.SHA256(),
        )
    except InvalidSignature as exc:
        raise SigningError("Manifest signature is not valid") from exc

    payload = json.loads(payload_bytes.decode("utf-8"))
    if payload.get("product") != PRODUCT:
        raise SigningError("Manifest is for product %r" % payload.get("product"))
    return payload


# ---------------------------------------------------------------------- CLI


def _cmd_genkey(args):
    hashes, serialization, padding, rsa = _require_cryptography()

    out = os.path.abspath(args.out)
    # A signing key committed by accident is the one mistake this system cannot
    # recover from without shipping a new client, so refuse the repository.
    if os.path.commonpath([out, ROOT]) == ROOT:
        raise SigningError(
            "Refusing to write the signing key inside the repository.\n"
            "Pick a path on offline or removable storage, for example:\n"
            "    python scripts/update_signing.py genkey --out D:/offline/deepexcel-update.pem"
        )
    if os.path.exists(out) and not args.force:
        raise SigningError("%s already exists; pass --force to overwrite" % out)

    key = rsa.generate_private_key(public_exponent=PUBLIC_EXPONENT, key_size=KEY_BITS)
    pem = key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    descriptor = os.open(out, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, "wb") as stream:
        stream.write(pem)

    modulus, exponent = public_key_bytes(key.public_key().public_numbers())
    modulus_b64 = base64.b64encode(modulus).decode("ascii")
    exponent_b64 = base64.b64encode(exponent).decode("ascii")
    write_embedded_key(modulus_b64, exponent_b64)

    print("[genkey] Private key: %s (%d-bit)" % (out, KEY_BITS))
    print("[genkey] Key id:      %s" % compute_key_id(modulus, exponent))
    print("[genkey] Public key written into %s" % os.path.relpath(SIGNING_CS, ROOT))
    print("")
    print("Move the private key to offline storage now. Anyone holding it can")
    print("ship an update that every installed client will accept and run.")
    print("Rebuild and re-release the client after this change: clients built")
    print("before it will reject manifests signed with this key.")


def _cmd_pubkey(args):
    info = describe_key(args.key)
    print(json.dumps(info, indent=2))


def _cmd_sign(args):
    with open(args.payload, "r", encoding="utf-8") as stream:
        payload = json.load(stream)
    manifest = sign_payload(args.key, payload)
    text = json.dumps(manifest, indent=2, ensure_ascii=False) + "\n"
    if args.out:
        with open(args.out, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(text)
        print("[sign] %s" % args.out)
    else:
        sys.stdout.write(text)


def _cmd_verify(args):
    embedded = read_embedded_key()
    if embedded is None:
        raise SigningError(
            "No public key is compiled into the client yet. Run:\n"
            "    python scripts/update_signing.py genkey --out <offline path>"
        )
    with open(args.manifest, "r", encoding="utf-8") as stream:
        manifest = json.load(stream)
    payload = verify_manifest(manifest, embedded["modulus_b64"], embedded["exponent_b64"])
    print("[verify] OK  key_id=%s version=%s channel=%s"
          % (manifest["key_id"], payload.get("version"), payload.get("channel")))


def build_test_vectors():
    """Cross-language test vectors for src/DeepExcel.Tests/UpdateManifestVectorTests.cs.

    The C# unit tests can sign with RSACng on their own, which proves the
    verifier is self-consistent but proves nothing about this signer. These
    vectors are produced by the real Python signing path and checked into the
    repository, so the C# side actually validates a manifest the release
    tooling would ship -- the PSS salt-length trap included.

    The signing key is generated here and discarded; only its public half is
    written out. Every rejection case below is a correctly signed manifest, so
    what is being tested is that authenticity alone is not enough.
    """
    hashes, serialization, padding, rsa = _require_cryptography()
    import tempfile

    key = rsa.generate_private_key(public_exponent=PUBLIC_EXPONENT, key_size=2048)
    modulus, exponent = public_key_bytes(key.public_key().public_numbers())
    handle, key_path = tempfile.mkstemp(suffix=".pem")
    try:
        with os.fdopen(handle, "wb") as stream:
            stream.write(key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption(),
            ))

        digest = "a" * 64
        base = dict(
            version="0.6.0", url="https://updates.example.com/DeepExcel.Setup.exe",
            sha256=digest, size=48 * 1024 * 1024, channel="stable",
            released_at="2026-09-13T00:00:00Z", notes="test vector",
        )

        def signed(**overrides):
            return sign_payload(key_path, build_payload(**dict(base, **overrides)))

        def signed_raw(mutate):
            payload = build_payload(**base)
            mutate(payload)
            return sign_payload(key_path, payload)

        cases = []

        cases.append({"name": "valid", "expect": "None", "expect_version": "0.6.0",
                      "manifest": signed()})

        cases.append({"name": "newer_patch", "expect": "None", "expect_version": "0.5.1",
                      "manifest": signed(version="0.5.1")})

        tampered = signed()
        altered = json.loads(base64.b64decode(tampered["payload"]).decode("utf-8"))
        altered["url"] = "https://evil.example.com/DeepExcel.Setup.exe"
        tampered["payload"] = base64.b64encode(
            json.dumps(altered, sort_keys=True, separators=(",", ":")).encode("utf-8")
        ).decode("ascii")
        cases.append({"name": "tampered_payload", "expect": "BadSignature", "manifest": tampered})

        flipped = signed()
        raw_signature = bytearray(base64.b64decode(flipped["signature"]))
        raw_signature[-1] ^= 0x01
        flipped["signature"] = base64.b64encode(bytes(raw_signature)).decode("ascii")
        cases.append({"name": "tampered_signature", "expect": "BadSignature", "manifest": flipped})

        foreign = signed()
        foreign["key_id"] = "0123456789abcdef"
        cases.append({"name": "unknown_key", "expect": "UnknownKey", "manifest": foreign})

        cases.append({"name": "downgrade", "expect": "NotNewer",
                      "manifest": signed(version="0.4.9")})
        cases.append({"name": "same_version", "expect": "NotNewer",
                      "manifest": signed(version="0.5.0")})
        cases.append({"name": "too_old_to_upgrade", "expect": "ManualUpgradeRequired",
                      "manifest": signed(minimum_upgradable_version="0.5.1")})
        cases.append({"name": "wrong_channel", "expect": "WrongChannel",
                      "manifest": signed(channel="beta")})

        cases.append({"name": "insecure_url", "expect": "InsecureUrl",
                      "manifest": signed_raw(
                          lambda p: p.__setitem__("url", "http://updates.example.com/s.exe"))})
        cases.append({"name": "credentials_in_url", "expect": "InsecureUrl",
                      "manifest": signed_raw(
                          lambda p: p.__setitem__("url", "https://u:p@updates.example.com/s.exe"))})
        cases.append({"name": "wrong_product", "expect": "WrongProduct",
                      "manifest": signed_raw(lambda p: p.__setitem__("product", "SomethingElse"))})
        cases.append({"name": "bad_digest", "expect": "BadDigest",
                      "manifest": signed_raw(lambda p: p.__setitem__("sha256", "zz" + "a" * 62))})
        cases.append({"name": "bad_size", "expect": "BadSize",
                      "manifest": signed_raw(lambda p: p.__setitem__("size", 0))})
        cases.append({"name": "oversized", "expect": "BadSize",
                      "manifest": signed_raw(lambda p: p.__setitem__("size", 2 * 1024 * 1024 * 1024))})
        cases.append({"name": "future_schema", "expect": "UnsupportedSchema",
                      "manifest": signed_raw(lambda p: p.__setitem__("schema", 2))})

        return {
            "comment": (
                "Generated by scripts/update_signing.py testvectors. Regenerate after "
                "changing the manifest format; the private key is ephemeral and discarded."
            ),
            "installed_version": "0.5.0",
            "channel": "stable",
            "key": {
                "modulus_b64": base64.b64encode(modulus).decode("ascii"),
                "exponent_b64": base64.b64encode(exponent).decode("ascii"),
                "key_id": compute_key_id(modulus, exponent),
            },
            "cases": cases,
        }
    finally:
        try:
            os.remove(key_path)
        except OSError:
            pass


def _cmd_testvectors(args):
    vectors = build_test_vectors()
    text = json.dumps(vectors, indent=2, ensure_ascii=False) + "\n"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(text)
    print("[testvectors] %s (%d cases)" % (args.out, len(vectors["cases"])))


def main(argv=None):
    parser = argparse.ArgumentParser(description="DeepExcel update manifest signing")
    sub = parser.add_subparsers(dest="command", required=True)

    genkey = sub.add_parser("genkey", help="Generate a signing key and embed its public half")
    genkey.add_argument("--out", required=True, help="Where to write the private key (outside this repo)")
    genkey.add_argument("--force", action="store_true")
    genkey.set_defaults(func=_cmd_genkey)

    pubkey = sub.add_parser("pubkey", help="Print the public half of a signing key")
    pubkey.add_argument("--key", required=True)
    pubkey.set_defaults(func=_cmd_pubkey)

    sign = sub.add_parser("sign", help="Sign a payload JSON file")
    sign.add_argument("--key", required=True)
    sign.add_argument("--payload", required=True)
    sign.add_argument("--out")
    sign.set_defaults(func=_cmd_sign)

    verify = sub.add_parser("verify", help="Verify a manifest against the compiled-in public key")
    verify.add_argument("--manifest", required=True)
    verify.set_defaults(func=_cmd_verify)

    vectors = sub.add_parser("testvectors", help="Regenerate the cross-language C# test vectors")
    vectors.add_argument(
        "--out",
        default=os.path.join(ROOT, "src", "DeepExcel.Tests", "fixtures", "update-test-vectors.json"))
    vectors.set_defaults(func=_cmd_testvectors)

    args = parser.parse_args(argv)
    args.func(args)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SigningError as error:
        print("ERROR: %s" % error, file=sys.stderr)
        sys.exit(1)
