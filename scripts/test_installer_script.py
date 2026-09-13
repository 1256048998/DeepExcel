#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Static checks on the Inno Setup script.

Run:  python scripts/test_installer_script.py

This is not a substitute for compiling with ISCC -- it cannot catch a Pascal
Script type error. It exists because the installer is the one component with no
automated verification at all on a machine without Inno Setup, and the class of
bug it does catch is the expensive one: the script referencing a file that the
packaging step never puts in the payload. That failure appears only on a user's
machine, as a silent rollback or a missing ribbon tab.

Everything checked here was chosen because it has actually gone wrong in this
project, or is one rename away from going wrong.
"""

import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ISS = os.path.join(ROOT, "deploy", "DeepExcel.Setup.iss")
PACKAGER = os.path.join(ROOT, "scripts", "package_release.py")

FAILURES = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" -- {detail}" if detail else ""))
    if not condition:
        FAILURES.append(name)


def read(path):
    with open(path, encoding="utf-8") as stream:
        return stream.read()


def strip_comments(iss):
    """Removes Inno Setup comments so checks match code, not prose.

    The comments deliberately explain what was removed and why -- "replaces the
    two powershell.exe hops" -- so a naive text search finds the very thing it
    is checking is absent.
    """
    lines = []
    for line in iss.splitlines():
        stripped = line.strip()
        # `;` starts a comment in the directive sections, `//` in [Code].
        if stripped.startswith(";") or stripped.startswith("//"):
            continue
        # Trailing `//` comment inside [Code]. Only cut when it is not inside a
        # string literal, which for this file means no quote before it.
        marker = line.find("//")
        if marker > 0 and line.count('"', 0, marker) % 2 == 0 and "'" not in line[:marker]:
            line = line[:marker]
        lines.append(line)
    text = "\n".join(lines)
    # Brace comments in Pascal Script. Inno also uses braces for constants such
    # as {app}, so only multi-word brace blocks are treated as comments.
    return re.sub(r"\{[^}\n]*\s[^}\n]*\}", " ", text)


def packaged_items(source, name):
    """Extracts a list literal from package_release.py."""
    match = re.search(name + r"\s*=\s*\[(.*?)\n\]", source, re.S)
    if not match:
        return set()
    items = set()
    for raw in re.findall(r'"([^"]+)"', match.group(1)):
        items.add(raw)
    # os.path.join("a", "b") entries contribute their first segment too.
    return items


def main():
    iss_raw = read(ISS)
    # Checks run against code with comments removed; the comments explain what
    # was deliberately taken out, so searching the raw text finds the very
    # strings being checked for absence.
    iss = strip_comments(iss_raw)
    packager = read(PACKAGER)

    print("=== installer script static checks ===")

    # ---- files the installer runs or links to must be packaged -----------
    shipped = packaged_items(packager, "EXCEL_ITEMS")
    required = packaged_items(packager, "EXCEL_REQUIRED_FILES")

    referenced = set()
    # ExpandConstant('{app}\Something.exe') and "{app}\Something.exe"
    for pattern in (r"\{app\}\\\\([A-Za-z0-9_.\\-]+)", r"\{app\}\\([A-Za-z0-9_.\-]+)"):
        referenced.update(re.findall(pattern, iss))
    referenced = {name for name in referenced if "." in name}

    check("installer references at least one packaged file", bool(referenced),
          f"found: {sorted(referenced)}")

    for name in sorted(referenced):
        check(f"{name} is in EXCEL_ITEMS", name in shipped,
              "installer would reference a file the packager never copies")
        check(f"{name} is in EXCEL_REQUIRED_FILES", name in required,
              "a missing file would produce a broken installer instead of a build failure")

    # ---- the powershell hops are gone -----------------------------------
    # They were replaced by native executables; PowerShell in the install path
    # is both a hard dependency on the user's environment and an antivirus
    # heuristic trigger.
    check("no powershell.exe in the install path",
          "powershell" not in iss.lower(),
          "PowerShell was deliberately removed from RunAndRequire")
    check("register-user.ps1 is not invoked",
          not re.search(r"RunAndRequire\([^)]*register-user", iss, re.S))
    check("PowerShellPath helper is gone", "function PowerShellPath" not in iss)

    # ---- both bitnesses are still registered -----------------------------
    # Spelling a WOW6432Node path from a 64-bit process does NOT reach the
    # 32-bit view; that mistake produced REGDB_E_CLASSNOTREG for several
    # releases.
    check("registers in the 64-bit view", "HKCU64" in iss)
    check("registers in the 32-bit view", "HKCU32" in iss)
    check("no hand-spelled WOW6432Node under HKCU64 for registration",
          not re.search(r"RegisterComClassView\(\s*HKCU64[^)]*WOW6432Node", iss, re.S))

    # ---- the 32-bit probe must actually be 32-bit ------------------------
    check("32-bit activation is verified by the x86 probe",
          "DeepExcel.Probe32.exe" in iss and "--activation-only" in iss)
    check("packager enforces the probe is a 32-bit image",
          "verify_probe_is_32bit" in packager)

    # ---- signing ---------------------------------------------------------
    check("self-signed thumbprint path is gone from the installer",
          "DEEPEXCEL_CERT_THUMBPRINT" not in iss_raw,
          "the internal self-signed build was removed")
    check("signing stays wired for a future certificate",
          "SignTool=deepsign" in iss_raw and "DEEPEXCEL_PFX" in iss_raw)

    # ---- rollback safety --------------------------------------------------
    check("a failed install restores the previous registration",
          "RestorePreviousExcelRegistration" in iss
          and "SnapshotPreviousExcelRegistration" in iss)
    # WPS activation used to be the last statement of the Excel try block, with
    # a revert if a later step failed. It is now its own try block, so:
    #   * an Excel-side failure no longer skips it -- the WPS add-in is plain JS
    #     and needs nothing from Excel, but it was unreachable on any machine
    #     where the Excel half failed (reported from the field, WPS installed
    #     but unusable), and on machines with no Excel at all;
    #   * nothing runs after it, so there is no "later step" left to revert for.
    # What must not regress is the independence itself.
    check("WPS activation is not chained behind the Excel block",
          re.search(r"Log\('Excel post-install step failed.*?\n\s*end;\s*\n.*?try\s*\n.*?UpdateWpsManifest\('Install'\)",
                    iss, re.S) is not None,
          "an Excel failure must not skip WPS activation")
    check("a WPS activation failure still fails the install",
          re.search(r"UpdateWpsManifest\('Install'\);.*?except.*?PostInstallFailed := True", iss, re.S)
          is not None)

    # ---- uninstall --------------------------------------------------------
    check("uninstall removes COM registration",
          "UnregisterDeepExcel" in iss and "CurUninstallStepChanged" in iss)

    # ---- version drift ----------------------------------------------------
    # The hardcoded 0.2.4.0 that broke v0.3-v0.4 is why this is generated.
    check("assembly version comes from the generated include",
          '#include "DeepExcel.reg.iss"' in iss_raw and "{#AssemblyVersion}" in iss_raw)
    check("packager generates that include", "generate_reg_iss" in packager)

    # ---- a failed install must not report success --------------------------
    # Inno treats an exception raised from CurStepChanged as non-fatal: it logs
    # the exception, continues to ssDone, and Setup still exits 0. Confirmed
    # with a minimal probe .iss. That is how a fresh-machine install could roll
    # back the COM registration and still report success -- the v0.4.11..v0.4.17
    # failure mode, hidden behind a zero exit code. Silent/IT deployments and CI
    # only ever see the exit code.
    check("post-install failure is recorded rather than re-raised",
          "PostInstallFailed := True" in iss and "RaiseException(errorMessage)" not in iss,
          "re-raising is silently swallowed by Inno and yields exit code 0")
    check("post-install failure forces a non-zero exit code",
          "ExitProcess@kernel32.dll" in iss and "ExitProcess(4)" in iss)
    check("the failure exit happens after rollback, in ssDone",
          iss.index("PostInstallFailed := True") < iss.index("ExitProcess(4)"))

    # ---- balanced Pascal blocks (cheap syntax smoke test) -----------------
    code = iss[iss.index("[Code]"):] if "[Code]" in iss else ""
    begins = len(re.findall(r"\bbegin\b", code))
    ends = len(re.findall(r"\bend[;.]", code))
    check("begin/end are balanced in [Code]", begins == ends,
          f"begin={begins} end={ends} -- not proof of correctness, only of gross imbalance")

    print()
    if FAILURES:
        print(f"FAILED ({len(FAILURES)}): {', '.join(FAILURES)}")
        return 1
    print("All checks passed.")
    print()
    print("NOTE: these are static checks -- they read the .iss, they do not run it.")
    print("Real verification is two more steps:")
    print("  python scripts/package_release.py --version <v>   (compiles the installer)")
    print("  powershell -File scripts/verify-install-sandbox.ps1 -Launch  (clean machine)")
    print("Inno Setup 6 lives in %LOCALAPPDATA%\\Programs\\Inno Setup 6, which is not")
    print("on PATH -- `where ISCC` finds nothing, so do not read that as 'not installed'.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
