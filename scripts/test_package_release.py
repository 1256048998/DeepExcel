#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for the release packaging guards.

Run:  python scripts/test_package_release.py

These cover the two guards whose failure modes are invisible at build time and
only surface as a broken install on a user's machine:

  * verify_probe_is_32bit -- an AnyCPU Probe32 would run 64-bit, test the wrong
    registry view, and report a false pass for 32-bit Office support.
  * write_checksums -- the integrity story for an unsigned build. If the digest
    file is wrong or missing, users have no way to verify a download.
"""

import hashlib
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSC = os.path.join(ROOT, "packages", "Microsoft.Net.Compilers.3.8.0", "tools", "csc.exe")
FRAMEWORK = r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319"


def load_module():
    path = os.path.join(ROOT, "scripts", "package_release.py")
    spec = importlib.util.spec_from_file_location("package_release", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


FAILURES = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print("[%s] %s%s" % (status, name, (" -- " + detail) if detail else ""))
    if not condition:
        FAILURES.append(name)


def expect_failure(name, callable_, needle):
    try:
        callable_()
    except Exception as exc:
        check(name, needle in str(exc), "raised: %s" % str(exc)[:70])
        return
    check(name, False, "expected a failure but none was raised")


def build_probe(platform, output):
    """Compile the real Probe32 sources at a chosen platform."""
    sources = [
        os.path.join(ROOT, "src", "DeepExcel.Probe32", "Program.cs"),
        os.path.join(ROOT, "src", "DeepExcel.Repair", "ComProbe.cs"),
        os.path.join(ROOT, "src", "DeepExcel.Repair", "AddInIdentity.cs"),
    ]
    command = [
        CSC, "/nologo", "/target:exe", "/platform:" + platform, "/out:" + output,
        "/reference:" + os.path.join(FRAMEWORK, "mscorlib.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.dll"),
        "/reference:" + os.path.join(FRAMEWORK, "System.Core.dll"),
    ] + sources
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError("csc failed for %s: %s" % (platform, result.stdout or result.stderr))
    return output


def test_probe_bitness_guard(module):
    if not os.path.isfile(CSC):
        print("[SKIP] probe bitness guard -- Roslyn compiler not available")
        return

    workspace = tempfile.mkdtemp(prefix="deepexcel-probe-test-")
    try:
        x86 = build_probe("x86", os.path.join(workspace, "x86.exe"))
        module.verify_probe_is_32bit(x86)
        check("x86 probe is accepted", True)

        # The whole point: an AnyCPU managed assembly has the SAME COFF machine
        # field as x86 (I386), so only the CLI header flag distinguishes them.
        anycpu = build_probe("anycpu", os.path.join(workspace, "anycpu.exe"))
        expect_failure("AnyCPU probe is rejected",
                       lambda: module.verify_probe_is_32bit(anycpu), "AnyCPU")

        x64 = build_probe("x64", os.path.join(workspace, "x64.exe"))
        expect_failure("x64 probe is rejected",
                       lambda: module.verify_probe_is_32bit(x64), "64-bit")

        not_pe = os.path.join(workspace, "garbage.exe")
        with open(not_pe, "wb") as stream:
            stream.write(b"not a windows executable at all")
        expect_failure("non-PE input is rejected",
                       lambda: module.verify_probe_is_32bit(not_pe), "not a Windows executable")
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def test_write_checksums(module):
    workspace = tempfile.mkdtemp(prefix="deepexcel-sums-test-")
    original_dist = module.DIST
    try:
        module.DIST = workspace
        payloads = {"DeepExcel.Setup.exe": b"installer bytes", "DeepExcel-v9.9.9.zip": b"zip bytes"}
        paths = []
        for name, content in payloads.items():
            path = os.path.join(workspace, name)
            with open(path, "wb") as stream:
                stream.write(content)
            paths.append(path)

        checksum_path = module.write_checksums(paths)
        with open(checksum_path, "r", encoding="ascii") as stream:
            lines = [line for line in stream.read().splitlines() if line]

        check("one digest line per artifact", len(lines) == len(payloads), "got %d" % len(lines))

        parsed = {}
        for line in lines:
            digest, name = line.split("  ", 1)
            parsed[name] = digest

        for name, content in payloads.items():
            expected = hashlib.sha256(content).hexdigest()
            check("digest correct for %s" % name, parsed.get(name) == expected)

        # sha256sum-compatible: "<64 hex>  <name>" and LF endings, so a user can
        # verify with standard tooling rather than trusting our own.
        with open(checksum_path, "rb") as stream:
            raw = stream.read()
        check("LF line endings", b"\r\n" not in raw)
        check("64-character hex digests", all(len(d) == 64 and all(c in "0123456789abcdef" for c in d)
                                              for d in parsed.values()))

        expect_failure("empty artifact list is rejected",
                       lambda: module.write_checksums([]), "No artifacts")
    finally:
        module.DIST = original_dist
        shutil.rmtree(workspace, ignore_errors=True)


def test_signing_rejects_self_signed_sources(module):
    saved = {name: os.environ.get(name) for name in
             ("DEEPEXCEL_PFX", "DEEPEXCEL_CERT_THUMBPRINT", "USE_AZURE_TRUSTED_SIGNING")}
    try:
        for name in saved:
            os.environ.pop(name, None)

        check("unsigned shell reports signing not configured", not module.signing_is_configured())

        os.environ["DEEPEXCEL_CERT_THUMBPRINT"] = "A" * 40
        check("thumbprint counts as a signing source", module.signing_is_configured())

        # Two sources at once is ambiguous and must never silently pick one.
        os.environ["USE_AZURE_TRUSTED_SIGNING"] = "1"
        expect_failure("two signing sources are rejected",
                       module.validate_production_signing_configuration, "exactly one signing source")

        os.environ.pop("USE_AZURE_TRUSTED_SIGNING", None)
        os.environ["DEEPEXCEL_CERT_THUMBPRINT"] = "not-a-thumbprint"
        expect_failure("malformed thumbprint is rejected",
                       module.validate_production_signing_configuration, "40-character")
    finally:
        for name, value in saved.items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value


def main():
    module = load_module()
    print("=== package_release guards ===")
    test_probe_bitness_guard(module)
    test_write_checksums(module)
    test_signing_rejects_self_signed_sources(module)

    print()
    if FAILURES:
        print("FAILED (%d): %s" % (len(FAILURES), ", ".join(FAILURES)))
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
