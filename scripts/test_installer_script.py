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
    # WPS has its own try block after the verification, so retiring the old JS
    # add-in always runs. Its add-in entry is written only when the COM classes
    # it loads survived verification: an entry for classes that were just rolled
    # back is a DeepExcel item in WPS that can never start.
    check("WPS handling is its own block after the verification",
          re.search(r"Log\('Post-install verification failed.*?\n\s*end;\s*\n.*?try\s*\n.*?RegisterWpsComAddIn\(",
                    iss, re.S) is not None,
          "retiring the JS add-in must not depend on the verification")
    check("WPS is registered only when the add-in verified",
          re.search(r"if not PostInstallFailed then\s*begin\s*RegisterWpsComAddIn\(", iss) is not None,
          "an entry for rolled-back classes cannot start")
    check("a WPS activation failure still fails the install",
          re.search(r"RegisterWpsComAddIn\([^;]*\);.*?except.*?PostInstallFailed := True", iss, re.S)
          is not None)

    # ---- WPS runs the Excel COM add-in (2026-11-23, real WPS 12.1) ----------
    # The old JS add-in cannot run in WPS (no Node in its add-in host); the COM
    # add-in does. Install registers COM for WPS and removes whatever an earlier
    # version left of the JS add-in, so a second, dead DeepExcel tab never
    # appears next to the working one.
    check("install registers the COM add-in for WPS",
          "'Software\\Kingsoft\\Office\\ET\\AddIns\\' + progId" in iss
          and "RegWriteDWordValue(HKCU, addinKey, 'LoadBehavior', 3)" in iss)
    check("the WPS whitelist entry is a value, not a subkey",
          "RegWriteStringValue(HKCU, 'Software\\Kingsoft\\Office\\ET\\AddinsWL', progId, '')" in iss,
          "AddinsWL lists ProgIDs as values; a subkey is ignored")
    check("install clears DeepExcel from WPS's disabled list",
          "RegDeleteValue(HKCU, 'Software\\Kingsoft\\Office\\ET\\AddinsCL', progId)" in iss)
    check("install ships no WPS JS add-in",
          "jsaddons\\DeepExcel_{#AppVersion}" not in iss and "{#PayloadDir}Wps" not in iss,
          "the JS add-in is retired; WPS loads the COM add-in")
    check("install retires the old JS add-in before registering COM for WPS",
          re.search(r"try\s*RetireWpsJsAddIn\(\);\s*except.*?end;\s*if not PostInstallFailed then",
                    iss, re.S) is not None,
          "a JS entry in publish.xml gives WPS a second, dead DeepExcel tab")
    check("a broken publish.xml does not fail the install",
          re.search(r"try\s*RetireWpsJsAddIn\(\);\s*except\s*Log\(", iss) is not None,
          "the cleanup is not what decides whether DeepExcel installed")
    check("retiring removes the publish.xml entry before the folders",
          re.search(r"procedure RetireWpsJsAddIn\(\);.*?RemoveWpsJsAddInEntry\(addinRoot\);"
                    r"\s*RemoveWpsJsAddInFolders\(addinRoot\);", iss, re.S) is not None,
          "the other order can leave an entry pointing at a deleted folder")
    check("retiring never creates a publish.xml",
          "ForceDirectories" not in iss[iss.find("procedure RemoveWpsJsAddInEntry"):iss.find("procedure RetireWpsJsAddIn")]
          and "loadXML" not in iss,
          "a machine that never had the JS add-in must be left alone")
    check("retiring removes only DeepExcel's own manifest nodes",
          "selectNodes('*[@name=\"DeepExcel\"]')" in iss)
    check("uninstall also retires the old JS add-in, and survives a broken manifest",
          re.search(r"usUninstall then.*?try\s*RetireWpsJsAddIn\(\);\s*except", iss, re.S) is not None)
    check("nothing re-registers rolled-back classes behind the verification's back",
          "ExcelNotActivated" not in iss,
          "the old WPS-only exemption re-registered classes that had just failed")
    check("uninstall removes the WPS COM registration",
          re.search(r"usUninstall then.*?UnregisterWpsComAddIn\('\{#MainProgId\}'\)", iss, re.S) is not None)

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
    # ---- silent runs must never wait for a click ------------------------
    # The auto-updater runs this installer with /VERYSILENT after Excel exits.
    # A message box there is not a cosmetic problem: nobody is present to
    # dismiss it, so the update hangs until the updater's timeout and is then
    # reported as a failure with no visible cause. WizardSilent alone was
    # observed returning False under /VERYSILENT, hence IsSilentRun.
    check("IsSilentRun exists", "function IsSilentRun()" in iss)
    check("IsSilentRun also reads the command line",
          "/VERYSILENT" in iss and "ParamStr(i)" in iss,
          "WizardSilent alone was unreliable; the command line is the fact")

    silent_guards = iss.count("IsSilentRun()")
    check("every dialog gate goes through IsSilentRun", silent_guards >= 4,
          f"found {silent_guards} uses (declaration + call sites)")

    # Exactly one: the fallback inside IsSilentRun itself.
    bare = len(re.findall(r"(?<!:= )WizardSilent", iss))
    check("no dialog is gated on WizardSilent alone", bare == 0,
          "a bare WizardSilent is the unreliable check this replaced")

    check("post-install failure is recorded rather than re-raised",
          "PostInstallFailed := True" in iss and "RaiseException(errorMessage)" not in iss,
          "re-raising is silently swallowed by Inno and yields exit code 0")
    check("post-install failure forces a non-zero exit code",
          "ExitProcess@kernel32.dll" in iss and "ExitProcess(4)" in iss)
    check("the failure exit happens after rollback, in ssDone",
          iss.index("PostInstallFailed := True") < iss.index("ExitProcess(4)"))
    # Activation needs .NET and the registration, not Excel. The old exemption
    # ("a WPS-only machine fails by definition") is what turned the elevated-host
    # failure of 2026-11-24 into a "successful" install whose add-in could never
    # load, so a failed verification is fatal on every machine.
    check("a failed verification is fatal with or without Excel",
          re.search(r"except\s*errorMessage := GetExceptionMessage\(\);.*?"
                    r"UnregisterDeepExcel\(\);\s*end;\s*PostInstallFailed := True;", iss, re.S) is not None,
          "the failure must not depend on whether Excel was detected")
    check(".NET 4.8 is required on WPS-only machines too",
          "(ExcelInstalled or IsWpsInstalled()) and (not IsDotNet48Detected())" in iss,
          "WPS loads the same .NET add-in")

    # ---- elevated hosts (UAC off / built-in Administrator) -----------------
    # Windows ignores HKCU\Software\Classes for elevated COM clients. Where
    # UAC is off every program is elevated, Excel and WPS included, so a
    # per-user-only registration is invisible to them (field case 2026-11-24:
    # 0x80070002; fixed on that machine by the same registration in HKLM).
    check("an elevated Setup also registers the classes machine-wide",
          "function MachineRegistrationNeeded" in iss and "Result := IsAdmin();" in iss
          and re.search(r"if MachineRegistrationNeeded\(\) then\s*begin.*?RegisterComClassView\(HKLM64.*?"
                        r"RegisterComClassView\(HKLM32", iss, re.S) is not None,
          "elevated Excel / WPS cannot see an HKCU-only registration")
    check("uninstall also removes the machine-wide classes",
          re.search(r"procedure UnregisterMachineClasses.*?RegDeleteKeyIncludingSubkeys\(HKLM64, 'SOFTWARE\\Classes\\CLSID\\'",
                    iss, re.S) is not None
          and re.search(r"procedure UnregisterComClass\(.*?\nbegin\s*UnregisterMachineClasses\(", iss, re.S) is not None)
    init = iss[iss.find("function InitializeSetup()"):iss.find("procedure CurStepChanged")]
    check("an elevated Setup replaces machine-wide residue without asking",
          re.search(r"HklmResidualExists\(\) and MachineRegistrationNeeded\(\).*?"
                    r"if HklmResidualExists\(\) then\s*begin\s*if IsSilentRun\(\) then", init, re.S) is not None,
          "otherwise the silent updater stops on a prompt it can never answer")
    check("a failed verification also drops the HKLM copy when restoring",
          re.search(r"RestorePreviousExcelRegistration\(\);\s*UnregisterMachineClasses\(", iss) is not None,
          "elevated hosts would keep loading the version that just failed")
    check("WebView2 is installed whatever host was detected",
          re.search(r"if WebView2NeedsInstall\(\) then\s*RunAndRequire", iss) is not None,
          "the verification treats a missing runtime as blocking on every machine")
    check("Excel is still always registered regardless of detection",
          "ExcelEnabled := True;" in iss)
    # WPS's startup guard ("此电脑已禁用可能拖慢启动速度的第三方加载项") keeps a
    # correctly registered add-in off and out of its COM add-in list; nothing
    # the installer writes clears it, so both finish messages must say how.
    finish = re.findall(r"if ([^\n]*?) then\s*MsgBox\('DeepExcel 安装完成！'(.*?)mbInformation", iss, re.S)
    wps_finish = [body for cond, body in finish if "not IsWpsInstalled()" not in cond]
    check("every finish message a WPS machine can see says how to allow the add-in",
          len(finish) >= 3 and len(wps_finish) >= 2
          and all("WpsEnableHint()" in body for body in wps_finish)
          and "点击此处进行管理" in iss,
          "a WPS user whose guard is on sees 'installed' and no DeepExcel tab")
    check("Excel-only machines are not told about WPS",
          any("not IsWpsInstalled()" in cond and "WpsEnableHint()" not in body for cond, body in finish))
    check("the finish messages no longer name only the first-seen prompt",
          "点 DeepExcel.AddIn 旁的“启用”即可" not in iss)

    # ---- sandbox verification script stays parseable ----------------------
    # verify-install-sandbox.ps1 is UTF-8 without a BOM, and Windows PowerShell
    # 5.1 decodes such .ps1 files with the system ANSI codepage. On a CJK
    # codepage a non-ASCII string literal eats its own closing quote and the
    # whole script fails to parse -- inside the sandbox that means the
    # LogonCommand dies and the run produces no result at all. Keep it ASCII.
    sandbox_ps1 = os.path.join(ROOT, "scripts", "verify-install-sandbox.ps1")
    with open(sandbox_ps1, "rb") as handle:
        sandbox_bytes = handle.read()
    non_ascii = [
        number
        for number, line in enumerate(sandbox_bytes.split(b"\n"), 1)
        if any(byte > 127 for byte in line)
    ]
    check("sandbox verification script is ASCII-only",
          not non_ascii,
          f"non-ASCII on lines {non_ascii[:5]}" if non_ascii else "")

    # ---- the sandbox run can finish without a human -----------------------
    # The verdict is written to a file and then the script holds the window with
    # "Press Enter to close". That is right for someone watching and wrong for
    # everyone else: an unattended run never returns, and the sandbox window
    # outlives the run on somebody's desktop.
    sandbox = sandbox_bytes.decode("ascii")

    check("sandbox script accepts -NoPause", "[switch]$NoPause" in sandbox)
    check("the closing hold honours -NoPause",
          "if ((-not $NoPause) -and $env:COMPUTERNAME" in sandbox,
          "otherwise the flag exists and changes nothing")

    # The subtle one. The host builds a .wsb whose LogonCommand re-invokes this
    # same script inside the guest. A flag added to the host half but left out
    # of that command line is accepted, reported as set, and silently does
    # nothing -- the guest never hears about it.
    logon = next((line for line in sandbox.splitlines() if "<Command>" in line), "")
    check("-NoPause reaches the guest through LogonCommand",
          "$pauseArg" in logon,
          "a flag missing from the LogonCommand is accepted and ignored")

    # ---- the sandbox checks WPS the way WPS now loads DeepExcel ------------
    # WPS loads the COM add-in. A sandbox run that still asserted on a jsaddons
    # payload would fail on a correct install, and one that instrumented main.js
    # would be measuring something that no longer ships.
    check("sandbox checks the WPS COM registration",
          "AddIns\\DeepExcel.AddIn" in sandbox and "AddinsWL" in sandbox)
    check("sandbox checks no old JS add-in is left",
          "no DeepExcel entry in WPS publish.xml" in sandbox
          and "no old WPS JS add-in folder" in sandbox)
    check("sandbox no longer looks for the JS add-in payload",
          "main.js" not in sandbox and "ComWps" not in sandbox)
    check("the WPS prompt asks whether the panel opens",
          "Open panel" in sandbox,
          "a ribbon tab alone does not prove the pane and sidecar work")

    # ---- balanced Pascal blocks (cheap syntax smoke test) -----------------
    # Every `end` closes a begin, a try or a case. Counting only `end;` against
    # `begin` was off by one per `end else` and per try block, and balanced only
    # by coincidence. String literals are blanked first so a message containing
    # the word cannot count.
    code = iss[iss.index("[Code]"):] if "[Code]" in iss else ""
    code = re.sub(r"'[^'\n]*'", "''", code)
    openers = len(re.findall(r"\b(?:begin|try|case)\b", code))
    ends = len(re.findall(r"\bend\b", code))
    check("begin/try/case and end are balanced in [Code]", openers == ends,
          f"openers={openers} end={ends} -- not proof of correctness, only of gross imbalance")

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
