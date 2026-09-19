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
import json
import os
import re
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


def test_updater_is_packaged(module):
    """The updater must reach the install directory, or nothing self-updates.

    It is also on the required-files list, because a missing one produces an
    install that looks perfect and can never upgrade itself -- the same class of
    omission that shipped WPS without jsplugins.xml.
    """
    check("DeepExcel.Updater.exe is copied into the payload",
          "DeepExcel.Updater.exe" in module.EXCEL_ITEMS)
    check("DeepExcel.Updater.exe is a required file",
          "DeepExcel.Updater.exe" in module.EXCEL_REQUIRED_FILES)


def test_every_wps_source_file_is_packaged(module):
    """Every tracked WPS source file must reach the release payload.

    This is the guard that was missing when jsplugins.xml shipped absent. The
    development path (build-wps.ps1) had always copied it; the release path's
    WPS_ITEMS did not list it. The result on a user's machine was an add-in
    whose ribbon tab appeared and whose every JS callback was dead -- reproduced
    in a clean sandbox against real WPS, and matching a field report of "WPS
    installs but does not work".

    Adding that one filename to the list fixed that one file. It did nothing
    about the next one. So the list is checked against the source tree itself
    rather than against the other hand-maintained list, which could drift the
    same way.

    Git is the source of truth: it excludes web/ and sidecar/, which are build
    outputs rather than sources.
    """
    wps_root = os.path.join(ROOT, "src", "DeepExcel.Wps")
    try:
        tracked = subprocess.run(
            ["git", "ls-files", "--", "src/DeepExcel.Wps"],
            cwd=ROOT, capture_output=True, text=True, timeout=60,
            encoding="utf-8", errors="replace")
    except OSError as exc:
        check("git is available to enumerate WPS sources", False, str(exc)[:70])
        return
    if tracked.returncode != 0:
        check("git is available to enumerate WPS sources", False, tracked.stderr[:70])
        return

    top_level_files, top_level_dirs = set(), set()
    for line in tracked.stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        relative = os.path.relpath(line, "src/DeepExcel.Wps").replace("\\", "/")
        head, _, tail = relative.partition("/")
        if tail:
            top_level_dirs.add(head)
        else:
            top_level_files.add(head)

    check("found WPS sources to check", bool(top_level_files),
          "%d files, %d directories" % (len(top_level_files), len(top_level_dirs)))

    packaged = set(module.WPS_ITEMS)
    missing = sorted(top_level_files - packaged)
    check("every tracked WPS file is in WPS_ITEMS", not missing,
          "not packaged: %s" % ", ".join(missing) if missing else "")

    missing_dirs = sorted(top_level_dirs - packaged)
    check("every tracked WPS directory is in WPS_ITEMS", not missing_dirs,
          "not packaged: %s" % ", ".join(missing_dirs) if missing_dirs else "")

    # The reverse: a name in the list that no longer exists makes the packager
    # fail at release time with "Required release item is missing", which is a
    # worse moment to find out than now.
    ghosts = sorted(
        name for name in packaged
        if not os.path.exists(os.path.join(wps_root, name.replace("/", os.sep)))
        and name not in ("web", "sidecar")
    )
    check("WPS_ITEMS lists nothing that no longer exists", not ghosts,
          "stale entries: %s" % ", ".join(ghosts) if ghosts else "")

    # jsplugins.xml specifically: it is the add-in manifest that binds ribbon
    # buttons to their handlers, so losing it produces a tab that looks fine
    # and does nothing -- the failure mode least likely to be noticed in a
    # smoke test.
    check("the add-in manifest is a required file",
          "jsplugins.xml" in module.WPS_REQUIRED_FILES,
          "without it a missing manifest becomes a silent user-facing failure")


def test_release_doc_covers_every_build_step(module):
    """The release runbook must build everything the packager demands.

    DEPLOYMENT.md went months missing build-repair.ps1 while three required
    executables came only from it -- anyone following the runbook hit
    "Required release item is missing" and had to work out why. Fixing the
    document fixes that one omission; this stops the next one.

    The mapping is derived, not listed: for each required .exe, find which
    build script emits it, then require that script to appear in the runbook's
    build section. Nothing here is hand-maintained, so nothing here can drift.
    """
    doc_path = os.path.join(ROOT, "docs", "DEPLOYMENT.md")
    with open(doc_path, encoding="utf-8") as stream:
        doc = stream.read()

    marker = "## 生产构建"
    if marker not in doc:
        check("DEPLOYMENT.md has a build section", False, "heading moved or renamed")
        return
    # Up to the next top-level heading.
    section = doc.split(marker, 1)[1]
    section = section.split("\n## ", 1)[0]

    # Which script emits which executable, read off the compiler's own /out:
    # argument. An earlier version of this searched whole files for the name
    # and was fooled immediately: _compile_only.ps1 mentions
    # DeepExcel.Updater.exe in a comment explaining a reference, so the guard
    # blamed the wrong script and would have sent someone to the wrong place.
    scripts_dir = os.path.join(ROOT, "scripts")
    producers = {}
    for candidate in sorted(os.listdir(scripts_dir)):
        if not candidate.endswith(".ps1"):
            continue
        with open(os.path.join(scripts_dir, candidate), encoding="utf-8",
                  errors="replace") as stream:
            for line in stream:
                if "/out:" not in line:
                    continue
                match = re.search(r"([A-Za-z0-9_.]+\.exe)", line)
                if match:
                    producers.setdefault(match.group(1), candidate)

    check("build scripts declare their executables", bool(producers),
          ", ".join(f"{k} <- {v}" for k, v in sorted(producers.items())))

    # What the runbook actually *runs*, not what it mentions. Prose counts as a
    # mention: the section now explains that build-repair.ps1 was once missing,
    # so a plain substring search passes even with the command deleted. That is
    # the same trap that made this guard blame the wrong script a moment ago.
    invoked = set()
    for line in section.splitlines():
        for pattern in (r"-File\s+scripts[\\/]([A-Za-z0-9_.-]+\.ps1)",
                        r"python\s+scripts[\\/]([A-Za-z0-9_.-]+\.py)"):
            found = re.search(pattern, line)
            if found:
                invoked.add(found.group(1))
    check("the runbook contains runnable build commands", bool(invoked),
          ", ".join(sorted(invoked)))

    traced = 0
    for name in module.EXCEL_REQUIRED_FILES:
        if not isinstance(name, str) or not name.endswith(".exe"):
            continue
        exe = os.path.basename(name.replace("\\", "/"))
        producer = producers.get(exe)
        if producer is None:
            # python.exe is downloaded by package-python.ps1, not compiled, so
            # it has no /out:. Not a gap -- just outside what this can trace.
            print(f"[SKIP] {exe} is not produced by a compiler step")
            continue
        traced += 1
        check(f"DEPLOYMENT.md builds {exe} (via {producer})",
              producer in invoked,
              f"the runbook never runs {producer}, so packaging stops on {exe}")

    check("at least one executable was traced to the runbook", traced > 0)

    # The manifest is the other thing a release can silently omit. Unlike a
    # missing executable it does not fail the build -- it just means nobody
    # upgrades, which looks identical to nobody having upgraded yet.
    check("DEPLOYMENT.md tells the releaser to publish update.json",
          "update.json" in doc,
          "without it an updated client base silently stops receiving releases")


def test_update_manifest_guards(module):
    """The update manifest guards.

    Every failure here is invisible: a client that stops upgrading looks exactly
    like a client nobody has upgraded yet.
    """
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    import update_signing

    try:
        update_signing._require_cryptography()
    except Exception as exc:
        check("cryptography is available for update signing", False, str(exc)[:80])
        return

    workspace = tempfile.mkdtemp(prefix="deepexcel-update-guards-")
    saved_dist = module.DIST
    saved_reader = update_signing.read_embedded_key
    saved_env = {
        name: os.environ.get(name)
        for name in ("DEEPEXCEL_UPDATE_KEY", "DEEPEXCEL_UPDATE_BASE_URL",
                     "DEEPEXCEL_UPDATE_CHANNEL", "DEEPEXCEL_UPDATE_NOTES",
                     "DEEPEXCEL_UPDATE_MINIMUM")
    }
    try:
        module.DIST = workspace
        for name in saved_env:
            os.environ.pop(name, None)

        setup = os.path.join(workspace, "DeepExcel.Setup.exe")
        with open(setup, "wb") as stream:
            stream.write(b"pretend installer")

        def make_key(path):
            hashes, serialization, padding, rsa = update_signing._require_cryptography()
            key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
            with open(path, "wb") as stream:
                stream.write(key.private_bytes(
                    encoding=serialization.Encoding.PEM,
                    format=serialization.PrivateFormat.PKCS8,
                    encryption_algorithm=serialization.NoEncryption()))
            return update_signing.describe_key(path)

        release_key = os.path.join(workspace, "release.pem")
        other_key = os.path.join(workspace, "other.pem")
        release_info = make_key(release_key)
        make_key(other_key)

        # A client with no compiled-in public key cannot accept updates at all,
        # so skipping the manifest is the consistent outcome, not an error.
        update_signing.read_embedded_key = lambda *a, **k: None
        check("no compiled-in key skips the manifest instead of failing",
              module.write_update_manifest("0.6.0", setup) is None)

        update_signing.read_embedded_key = lambda *a, **k: release_info
        expect_failure(
            "compiled-in key with no signing key is a hard failure",
            lambda: module.write_update_manifest("0.6.0", setup),
            "DEEPEXCEL_UPDATE_KEY is not set")

        os.environ["DEEPEXCEL_UPDATE_KEY"] = release_key
        expect_failure(
            "a missing download base URL is a hard failure",
            lambda: module.write_update_manifest("0.6.0", setup),
            "DEEPEXCEL_UPDATE_BASE_URL is required")

        os.environ["DEEPEXCEL_UPDATE_BASE_URL"] = "http://updates.example.com/v0.6.0"
        expect_failure(
            "a plain-http download URL is rejected",
            lambda: module.write_update_manifest("0.6.0", setup),
            "must be https")

        os.environ["DEEPEXCEL_UPDATE_BASE_URL"] = "https://updates.example.com/v0.6.0"

        # The one that matters: signing with a key the shipped client does not
        # know breaks updating for everyone, all at once, with no symptom.
        os.environ["DEEPEXCEL_UPDATE_KEY"] = other_key
        expect_failure(
            "signing with a key the client does not know is a hard failure",
            lambda: module.write_update_manifest("0.6.0", setup),
            "would reject this release")

        os.environ["DEEPEXCEL_UPDATE_KEY"] = release_key
        path = module.write_update_manifest("0.6.0", setup)
        check("a matching key publishes a manifest", path is not None and os.path.isfile(path))

        with open(path, encoding="utf-8") as stream:
            manifest = json.load(stream)
        payload = update_signing.verify_manifest(
            manifest, release_info["modulus_b64"], release_info["exponent_b64"])
        check("the published manifest verifies under the client's key", True)
        check("the manifest carries the installer's real digest",
              payload["sha256"] == hashlib.sha256(b"pretend installer").hexdigest())
        check("the manifest carries the installer's real size",
              payload["size"] == len(b"pretend installer"))
        check("the download URL is the published one",
              payload["url"] == "https://updates.example.com/v0.6.0/DeepExcel.Setup.exe")
    finally:
        module.DIST = saved_dist
        update_signing.read_embedded_key = saved_reader
        for name, value in saved_env.items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value
        shutil.rmtree(workspace, ignore_errors=True)


def main():
    module = load_module()
    print("=== package_release guards ===")
    test_probe_bitness_guard(module)
    test_write_checksums(module)
    test_signing_rejects_self_signed_sources(module)
    test_updater_is_packaged(module)
    test_every_wps_source_file_is_packaged(module)
    test_release_doc_covers_every_build_step(module)
    test_update_manifest_guards(module)

    print()
    if FAILURES:
        print("FAILED (%d): %s" % (len(FAILURES), ", ".join(FAILURES)))
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
