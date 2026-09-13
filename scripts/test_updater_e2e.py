#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""End-to-end check of the update chain: real key, real signature, real binary.

The C# unit tests verify manifests in-process and the Python tests verify the
signer. Neither covers the thing that actually ships: DeepExcel.Updater.exe,
compiled against a compiled-in public key, reading a staged directory off disk
and deciding whether to run an installer.

This builds the updater against a throwaway signing key -- into a temporary
directory, so the repository's own (deliberately empty) key is never touched --
stages manifests and asserts the exit code for each scenario.

    python scripts/test_updater_e2e.py

Requires: cryptography, and packages/Microsoft.Net.Compilers (already vendored).
"""

import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import update_signing  # noqa: E402

CSC = os.path.join(ROOT, "packages", "Microsoft.Net.Compilers.3.8.0", "tools", "csc.exe")
FRAMEWORK = r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
UPDATER_DIR = os.path.join(ROOT, "src", "DeepExcel.Updater")
UPDATES_DIR = os.path.join(ROOT, "src", "DeepExcel.AddIn", "Updates")
SHARED_SOURCES = ["ReleaseVersion.cs", "UpdateSigning.cs", "UpdateManifest.cs", "UpdateStage.cs"]

EXIT_INSTALLED = 0
EXIT_USAGE = 2
EXIT_STAGE_REJECTED = 3
EXIT_DIGEST_MISMATCH = 4

_failures = []
_checks = 0


def check(condition, message):
    global _checks
    _checks += 1
    if not condition:
        _failures.append(message)
        print("  FAIL  %s" % message)
    else:
        print("  ok    %s" % message)


def build_updater(workspace, private_key_path):
    """Compiles DeepExcel.Updater.exe with the test key compiled in."""
    sources = os.path.join(workspace, "src")
    os.makedirs(sources)
    for name in SHARED_SOURCES:
        shutil.copy2(os.path.join(UPDATES_DIR, name), os.path.join(sources, name))
    shutil.copy2(os.path.join(UPDATER_DIR, "Program.cs"), os.path.join(sources, "Program.cs"))

    info = update_signing.describe_key(private_key_path)
    update_signing.write_embedded_key(
        info["modulus_b64"], info["exponent_b64"],
        source_path=os.path.join(sources, "UpdateSigning.cs"))

    output = os.path.join(workspace, "DeepExcel.Updater.exe")
    command = [
        CSC, "/target:winexe", "/out:" + output, "/langversion:9.0", "/nologo",
        "/platform:anycpu",
        "/win32manifest:" + os.path.join(UPDATER_DIR, "app.manifest"),
        "/reference:" + os.path.join(FRAMEWORK, "mscorlib.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.Core.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.Windows.Forms.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.Web.Extensions.dll"),
        os.path.join(sources, "Program.cs"),
    ] + [os.path.join(sources, name) for name in SHARED_SOURCES]

    result = subprocess.run(command, capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
    if result.returncode != 0:
        raise RuntimeError("Updater build failed:\n%s\n%s" % (result.stdout, result.stderr))
    return output, info


def stage(workspace, name, private_key_path, version="0.6.0", package=b"pretend installer bytes",
          channel="stable", minimum=None):
    directory = os.path.join(workspace, "stages", name, version)
    os.makedirs(directory)
    package_path = os.path.join(directory, "DeepExcel.Setup.exe")
    with open(package_path, "wb") as stream:
        stream.write(package)

    payload = update_signing.build_payload(
        version=version,
        url="https://updates.example.com/DeepExcel.Setup.exe",
        sha256=hashlib.sha256(package).hexdigest(),
        size=len(package),
        channel=channel,
        released_at="2026-09-13T00:00:00Z",
        minimum_upgradable_version=minimum,
    )
    manifest = update_signing.sign_payload(private_key_path, payload)
    with open(os.path.join(directory, "update.json"), "w", encoding="utf-8", newline="\n") as stream:
        json.dump(manifest, stream, indent=2)
    return directory


_slowest = 0.0


def run_updater(updater, stage_dir, installed="0.5.0", extra=None):
    global _slowest
    command = [updater, "--stage", stage_dir, "--installed-version", installed, "--dry-run"]
    if extra:
        command += extra
    started = time.monotonic()
    # A short timeout on purpose. The updater must never block waiting for
    # someone to click something unless --notify was passed, and a modal dialog
    # is exactly the kind of thing that would otherwise appear on a developer's
    # desktop in the middle of a test run.
    result = subprocess.run(command, capture_output=True, text=True, timeout=30,
                            encoding="utf-8", errors="replace")
    _slowest = max(_slowest, time.monotonic() - started)
    return result.returncode, (result.stdout or "") + (result.stderr or "")


def main():
    if not os.path.isfile(CSC):
        print("ERROR: Roslyn compiler not found: %s" % CSC, file=sys.stderr)
        return 1

    workspace = tempfile.mkdtemp(prefix="deepexcel-updater-e2e-")
    key_path = os.path.join(workspace, "test-signing-key.pem")
    try:
        hashes, serialization, padding, rsa = update_signing._require_cryptography()
        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        with open(key_path, "wb") as stream:
            stream.write(key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()))

        print("==> Building DeepExcel.Updater.exe with a throwaway signing key")
        updater, info = build_updater(workspace, key_path)
        print("    key_id %s  ->  %s" % (info["key_id"], updater))

        print("\n==> Scenarios")

        code, output = run_updater(updater, stage(workspace, "valid", key_path))
        check(code == EXIT_INSTALLED, "a correctly signed, newer release is accepted (exit %d)" % code)
        check("清单已验签" in output, "the log records manifest verification")
        check("更新包校验通过" in output, "the log records package verification")

        # The package is what actually gets executed, so a byte changed after
        # staging is the case that matters most.
        swapped = stage(workspace, "swapped", key_path)
        with open(os.path.join(swapped, "DeepExcel.Setup.exe"), "wb") as stream:
            stream.write(b"pretend installer bytes".replace(b"p", b"P"))
        code, _ = run_updater(updater, swapped)
        check(code == EXIT_DIGEST_MISMATCH,
              "a package swapped after staging is refused (exit %d)" % code)

        forged = stage(workspace, "forged", key_path)
        manifest_path = os.path.join(forged, "update.json")
        with open(manifest_path, encoding="utf-8") as stream:
            manifest = json.load(stream)
        raw = bytearray(base64.b64decode(manifest["signature"]))
        raw[-1] ^= 0x01
        manifest["signature"] = base64.b64encode(bytes(raw)).decode("ascii")
        with open(manifest_path, "w", encoding="utf-8") as stream:
            json.dump(manifest, stream)
        code, output = run_updater(updater, forged)
        check(code == EXIT_STAGE_REJECTED, "a tampered signature is refused (exit %d)" % code)
        check("BadSignature" in output, "the refusal names the reason")

        # Signed by a key this build does not know: the exact scenario where an
        # attacker controls the update server but not the signing key.
        other_key = os.path.join(workspace, "other-key.pem")
        other = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        with open(other_key, "wb") as stream:
            stream.write(other.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()))
        code, output = run_updater(updater, stage(workspace, "foreign", other_key))
        check(code == EXIT_STAGE_REJECTED, "a manifest from a foreign key is refused (exit %d)" % code)
        check("UnknownKey" in output, "a foreign key is reported as UnknownKey")

        code, _ = run_updater(updater, stage(workspace, "older", key_path, version="0.4.0"))
        check(code == EXIT_STAGE_REJECTED, "a downgrade is refused (exit %d)" % code)

        code, _ = run_updater(updater, stage(workspace, "beta", key_path, channel="beta"))
        check(code == EXIT_STAGE_REJECTED, "another channel's release is refused (exit %d)" % code)

        code, _ = run_updater(
            updater, stage(workspace, "floor", key_path, minimum="0.5.1"))
        check(code == EXIT_STAGE_REJECTED,
              "a build below the upgrade floor is refused (exit %d)" % code)

        empty = os.path.join(workspace, "stages", "empty")
        os.makedirs(empty)
        code, _ = run_updater(updater, empty)
        check(code == EXIT_STAGE_REJECTED, "an empty stage is refused (exit %d)" % code)

        result = subprocess.run([updater], capture_output=True, text=True, timeout=60,
                                encoding="utf-8", errors="replace")
        check(result.returncode == EXIT_USAGE,
              "no arguments is a usage error, not a crash (exit %d)" % result.returncode)

        log = os.path.join(stage(workspace, "logged", key_path), "updater.log")
        run_updater(updater, os.path.dirname(log))
        check(os.path.isfile(log), "the updater leaves a log in the stage directory")

        # Every scenario above failed in some way, and none was given --notify.
        # If a failure ever opens a dialog by default it will sit on somebody's
        # desktop waiting for a click, and the only symptom is that this gets
        # slow. Anything past a couple of seconds means it is waiting on a human.
        check(_slowest < 5.0,
              "a failure without --notify never waits for a click (slowest run %.1fs)" % _slowest)

        failing = stage(workspace, "notify-off", key_path)
        with open(os.path.join(failing, "DeepExcel.Setup.exe"), "wb") as stream:
            stream.write(b"corrupted")
        _, output = run_updater(updater, failing)
        check("更新包已损坏或被替换" in output,
              "a silent failure still reports the reason on stderr")
    finally:
        shutil.rmtree(workspace, ignore_errors=True)

    print("\n%d checks, %d failed" % (_checks, len(_failures)))
    return 1 if _failures else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:  # noqa: BLE001
        print("ERROR: %s" % error, file=sys.stderr)
        sys.exit(1)
