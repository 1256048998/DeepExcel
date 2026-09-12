#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a deterministic DeepExcel release payload and unified installer.

Usage:
    python scripts/package_release.py --version 0.4.17

The installer supports both Microsoft Excel (managed COM add-in) and WPS
Spreadsheets (WPS publish.xml add-in).  Missing runtime files are fatal: a
release build must never silently produce a package that cannot load.
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import urllib.request
import zipfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN_RELEASE = os.path.abspath(os.environ.get(
    "DEEPEXCEL_BIN_RELEASE",
    os.path.join(ROOT, "src", "DeepExcel.AddIn", "bin", "Release"),
))
WPS_SOURCE = os.path.join(ROOT, "src", "DeepExcel.Wps")
DIST = os.path.join(ROOT, "dist")
SCRIPTS = os.path.join(ROOT, "scripts")
DEPLOY = os.path.join(ROOT, "deploy")
PAYLOAD = os.path.join(DEPLOY, "payload")
WEBVIEW_BOOTSTRAPPER = os.path.join(DEPLOY, "Includes", "MicrosoftEdgeWebview2Setup.exe")
WEBVIEW_BOOTSTRAPPER_URL = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
INTERNAL_SIGNING_DIR = os.path.join(ROOT, ".local-signing")
INTERNAL_CERTIFICATE = os.path.join(INTERNAL_SIGNING_DIR, "DeepExcel.Internal.cer")
INTERNAL_THUMBPRINT = os.path.join(INTERNAL_SIGNING_DIR, "certificate-thumbprint.txt")

EXCEL_ITEMS = [
    "DeepExcel.AddIn.dll",
    "DeepExcel.AddIn.dll.config",
    "Extensibility.dll",
    "Microsoft.Bcl.AsyncInterfaces.dll",
    "Microsoft.Office.Interop.Excel.dll",
    "Microsoft.Vbe.Interop.dll",
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll",
    "OFFICE.dll",
    "System.Buffers.dll",
    "System.Memory.dll",
    "System.Numerics.Vectors.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Text.Encodings.Web.dll",
    "System.Text.Json.dll",
    "System.Threading.Tasks.Extensions.dll",
    "System.ValueTuple.dll",
    "WebViewAssets",
    "sidecar",
    "python",
    "runtimes",
]

EXCEL_REQUIRED_FILES = [
    "DeepExcel.AddIn.dll",
    "DeepExcel.AddIn.dll.config",
    "Extensibility.dll",
    "Microsoft.Office.Interop.Excel.dll",
    "Microsoft.Vbe.Interop.dll",
    "OFFICE.dll",
    os.path.join("WebViewAssets", "index.html"),
    os.path.join("WebViewAssets", "assets", "index.js"),
    os.path.join("WebViewAssets", "assets", "index.css"),
    os.path.join("sidecar", "sidecar.py"),
    os.path.join("python", "python.exe"),
]

WPS_ITEMS = [
    "main.js",
    "ribbon.xml",
    "taskpane.html",
    "package.json",
    "sidecar-host.js",
    "tool-dispatcher.js",
    "wps-actions.js",
    "jsa-executor.js",
    # 模型配置（厂商 / 模型优先级 / API Key），与 Excel 端共用 config.json + DPAPI 凭据
    "config-store.js",
    "credential-store.js",
    "model-service.js",
    "dpapi-cli.py",
    # 对话历史 + 附件
    "conversation-store.js",
    "attachment-store.js",
    "sidecar",
    "web",
    "images",
]

WPS_REQUIRED_FILES = [
    "main.js",
    "ribbon.xml",
    "taskpane.html",
    "sidecar-host.js",
    "config-store.js",
    "credential-store.js",
    "model-service.js",
    "dpapi-cli.py",
    "conversation-store.js",
    "attachment-store.js",
    os.path.join("sidecar", "sidecar.py"),
    os.path.join("web", "index.html"),
    os.path.join("web", "assets", "index.js"),
    os.path.join("web", "assets", "index.css"),
    os.path.join("images", "panel.svg"),
    os.path.join("images", "help.svg"),
]

INSTALL_TXT_TEMPLATE = """DeepExcel v{version}
=====================

Recommended installation
------------------------
Run DeepExcel.Setup.exe. The same installer enables DeepExcel for Microsoft
Excel and WPS Spreadsheets, installs the embedded Python runtime, and installs
the WebView2 runtime when required. Administrator access is not normally needed.

Before installing, close Excel and WPS Spreadsheets. After installation, reopen
either application and look for the DeepExcel ribbon tab.

The ZIP is a support/debug payload. End users should receive DeepExcel.Setup.exe.
"""


def fail(message):
    raise RuntimeError(message)


def safe_recreate_dir(path):
    expected = os.path.abspath(PAYLOAD)
    resolved = os.path.abspath(path)
    if resolved != expected:
        fail("Refusing to recreate unexpected payload path: %s" % resolved)
    if os.path.isdir(resolved):
        shutil.rmtree(resolved)
    os.makedirs(resolved)


def copy_item(source_root, item, destination_root):
    source = os.path.join(source_root, item)
    destination = os.path.join(destination_root, item)
    if not os.path.exists(source):
        fail("Required release item is missing: %s" % source)
    if os.path.isdir(source):
        shutil.copytree(
            source,
            destination,
            ignore=shutil.ignore_patterns("__pycache__", "tests", "*.pyc", "*.pyo"),
        )
    else:
        os.makedirs(os.path.dirname(destination), exist_ok=True)
        shutil.copy2(source, destination)


def read_assembly_version(dll_path):
    ps = (
        "[System.Reflection.AssemblyName]::GetAssemblyName("
        "'%s').Version.ToString()" % dll_path.replace("'", "''")
    )
    out = subprocess.check_output(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        stderr=subprocess.STDOUT,
        text=True,
        timeout=60,
    ).strip()
    if not out or not out[0].isdigit():
        fail("Cannot read assembly version from %s" % dll_path)
    return out


def generate_reg_iss(version):
    os.makedirs(DEPLOY, exist_ok=True)
    assembly = "DeepExcel.AddIn, Version=%s, Culture=neutral, PublicKeyToken=null" % version
    content = (
        "; Auto-generated by scripts/package_release.py -- DO NOT EDIT\n"
        "#define AssemblyVersion \"%s\"\n"
        "#define AssemblyValue \"%s\"\n" % (version, assembly)
    )
    path = os.path.join(DEPLOY, "DeepExcel.reg.iss")
    with open(path, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(content)
    return path


def sign_files(paths):
    sign_script = os.path.join(SCRIPTS, "sign.ps1")
    existing = [path for path in paths if os.path.isfile(path)]
    if not existing:
        return
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", sign_script] + existing,
        check=True,
    )


def signing_is_configured():
    return bool(os.environ.get("DEEPEXCEL_PFX")) or os.environ.get(
        "USE_AZURE_TRUSTED_SIGNING", ""
    ).lower() in ("1", "true") or bool(os.environ.get("DEEPEXCEL_CERT_THUMBPRINT"))


def validate_production_signing_configuration():
    """Return the expected PFX signer thumbprint, or None for Azure signing."""
    if os.environ.get("DEEPEXCEL_CERT_THUMBPRINT"):
        fail("DEEPEXCEL_CERT_THUMBPRINT is internal-only and is forbidden for production builds")

    pfx = os.environ.get("DEEPEXCEL_PFX")
    use_azure = os.environ.get("USE_AZURE_TRUSTED_SIGNING", "").lower() in ("1", "true")
    if bool(pfx) == bool(use_azure):
        fail("Production builds require exactly one signing source: PFX or Azure Trusted Signing")

    if use_azure:
        # Inno Setup's preprocessor is case-sensitive; normalize values such as
        # TRUE so it also signs the nested uninstaller.
        os.environ["USE_AZURE_TRUSTED_SIGNING"] = "1"
        return None

    if not os.path.isfile(pfx):
        fail("Production PFX does not exist: %s" % pfx)
    command = (
        "$c=[Security.Cryptography.X509Certificates.X509Certificate2]::new("
        "$env:DEEPEXCEL_PFX,$env:DEEPEXCEL_PFX_PASS,"
        "[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet);"
        "$eku=$c.Extensions|Where-Object{$_.Oid.Value -eq '2.5.29.37'}|"
        "ForEach-Object{$_.EnhancedKeyUsages}|Where-Object{$_.Value -eq '1.3.6.1.5.5.7.3.3'};"
        "if(-not $c.HasPrivateKey -or -not $eku -or $c.Subject -eq $c.Issuer "
        "-or $c.NotBefore -gt (Get-Date) -or $c.NotAfter -le (Get-Date)){exit 1};"
        "$chain=[Security.Cryptography.X509Certificates.X509Chain]::new();"
        "$chain.ChainPolicy.RevocationMode='NoCheck';"
        "if(-not $chain.Build($c)){exit 1};$c.Thumbprint"
    )
    try:
        thumbprint = subprocess.check_output(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=30,
        ).strip().upper()
    except subprocess.CalledProcessError as exc:
        fail(
            "Production PFX must contain a currently valid, CA-issued code-signing "
            "certificate with a private key: %s" % exc.output.strip()
        )
    if len(thumbprint) != 40 or any(ch not in "0123456789ABCDEF" for ch in thumbprint):
        fail("Cannot determine the production PFX signer thumbprint")
    return thumbprint


def load_internal_cert_thumbprint():
    if not os.path.isfile(INTERNAL_CERTIFICATE) or not os.path.isfile(INTERNAL_THUMBPRINT):
        fail(
            "Internal signing certificate is missing; run "
            "scripts/new-internal-signing-cert.ps1 first"
        )
    with open(INTERNAL_THUMBPRINT, "r", encoding="ascii") as stream:
        thumbprint = "".join(stream.read().split()).upper()
    if len(thumbprint) != 40 or any(ch not in "0123456789ABCDEF" for ch in thumbprint):
        fail("Invalid internal signing certificate thumbprint")
    escaped_certificate = INTERNAL_CERTIFICATE.replace("'", "''")
    command = (
        "$c=[Security.Cryptography.X509Certificates.X509Certificate2]::new('%s');"
        "$eku=$c.Extensions|Where-Object{$_.Oid.Value -eq '2.5.29.37'}|"
        "ForEach-Object{$_.EnhancedKeyUsages}|Where-Object{$_.Value -eq '1.3.6.1.5.5.7.3.3'};"
        "if($c.Subject -ne 'CN=DeepExcel Internal Testing' -or $c.Issuer -ne $c.Subject "
        "-or -not $eku -or $c.NotBefore -gt (Get-Date) -or $c.NotAfter -le (Get-Date)){exit 1};"
        "$c.Thumbprint" % escaped_certificate
    )
    certificate_thumbprint = subprocess.check_output(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
        stderr=subprocess.STDOUT,
        text=True,
        timeout=30,
    ).strip().upper()
    if certificate_thumbprint != thumbprint:
        fail("Internal CER and certificate-thumbprint.txt do not match")
    return thumbprint


def verify_authenticode(path, publisher_pattern=None, expected_thumbprint=None):
    escaped_path = path.replace("'", "''")
    checks = ["$s.Status -ne 'Valid'"]
    if publisher_pattern:
        escaped_pattern = publisher_pattern.replace("'", "''")
        checks.append("$s.SignerCertificate.Subject -notmatch '%s'" % escaped_pattern)
    if expected_thumbprint:
        escaped_thumbprint = expected_thumbprint.replace("'", "''")
        checks.append("$s.SignerCertificate.Thumbprint -ne '%s'" % escaped_thumbprint)
    command = (
        "$m=Join-Path $env:WINDIR 'System32\\WindowsPowerShell\\v1.0\\Modules\\Microsoft.PowerShell.Security\\Microsoft.PowerShell.Security.psd1';"
        "Import-Module -Name $m -Force -ErrorAction Stop;"
        "$s=Get-AuthenticodeSignature -LiteralPath '%s';"
        "if(%s){Write-Error ('Invalid Authenticode signature: '+$s.Status+' '+$s.SignerCertificate.Subject);exit 1}"
        % (escaped_path, " -or ".join(checks))
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
        check=True,
    )


def ensure_webview2_bootstrapper():
    os.makedirs(os.path.dirname(WEBVIEW_BOOTSTRAPPER), exist_ok=True)
    if not os.path.isfile(WEBVIEW_BOOTSTRAPPER) or os.path.getsize(WEBVIEW_BOOTSTRAPPER) < 100_000:
        print("[runtime] Downloading Microsoft's WebView2 Evergreen bootstrapper...")
        temporary = WEBVIEW_BOOTSTRAPPER + ".download"
        urllib.request.urlretrieve(WEBVIEW_BOOTSTRAPPER_URL, temporary)
        with open(temporary, "rb") as stream:
            if stream.read(2) != b"MZ":
                os.remove(temporary)
                fail("Downloaded WebView2 bootstrapper is not a Windows executable")
        os.replace(temporary, WEBVIEW_BOOTSTRAPPER)

    verify_authenticode(WEBVIEW_BOOTSTRAPPER, "Microsoft")


def assemble_payload(version):
    safe_recreate_dir(PAYLOAD)
    excel_payload = os.path.join(PAYLOAD, "Excel")
    wps_payload = os.path.join(PAYLOAD, "Wps")
    os.makedirs(excel_payload)
    os.makedirs(wps_payload)

    for item in EXCEL_ITEMS:
        copy_item(BIN_RELEASE, item, excel_payload)
    for item in WPS_ITEMS:
        copy_item(WPS_SOURCE, item, wps_payload)

    for relative_path in WPS_REQUIRED_FILES:
        packaged_path = os.path.join(wps_payload, relative_path)
        if not os.path.isfile(packaged_path) or os.path.getsize(packaged_path) == 0:
            fail("Missing or empty WPS runtime asset: %s" % packaged_path)

    for relative_path in EXCEL_REQUIRED_FILES:
        packaged_path = os.path.join(excel_payload, relative_path)
        if not os.path.isfile(packaged_path) or os.path.getsize(packaged_path) == 0:
            fail("Missing or empty Excel runtime asset: %s" % packaged_path)

    # Cold-import with the exact packaged interpreter and its natural _pth
    # configuration. A manual sys.path insertion here previously allowed a
    # package that crashed immediately with "No module named excel_tools".
    packaged_python = os.path.join(excel_payload, "python", "python.exe")
    try:
        subprocess.run(
            [packaged_python, "-c", "import sidecar"],
            cwd=excel_payload,
            check=True,
            capture_output=True,
            text=True,
            timeout=30,
        )
    except subprocess.CalledProcessError as exc:
        fail("Packaged sidecar cold import failed: %s" % (exc.stderr or exc.stdout))

    with open(os.path.join(wps_payload, "package.json"), "r", encoding="utf-8") as stream:
        wps_package = json.load(stream)
    if wps_package.get("version") != version:
        fail(
            "WPS package version %r does not match release version %s"
            % (wps_package.get("version"), version)
        )

    for script_name in ("register-user.ps1", "diagnose.ps1"):
        copy_item(SCRIPTS, script_name, excel_payload)

    with open(os.path.join(PAYLOAD, "INSTALL.txt"), "w", encoding="utf-8") as stream:
        stream.write(INSTALL_TXT_TEMPLATE.format(version=version))

    if os.path.exists(os.path.join(excel_payload, "WebView2Loader.dll")):
        fail("A root-level WebView2Loader.dll is forbidden for AnyCPU builds")
    for arch in ("x86", "x64", "arm64"):
        loader = os.path.join(excel_payload, "runtimes", "win-%s" % arch, "native", "WebView2Loader.dll")
        if not os.path.isfile(loader):
            fail("Missing WebView2 native loader for %s: %s" % (arch, loader))


def create_zip(version):
    os.makedirs(DIST, exist_ok=True)
    path = os.path.join(DIST, "DeepExcel-v%s.zip" % version)
    if os.path.exists(path):
        os.remove(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for current_root, directories, files in os.walk(PAYLOAD):
            directories.sort()
            for filename in sorted(files):
                full_path = os.path.join(current_root, filename)
                archive_name = os.path.relpath(full_path, PAYLOAD).replace(os.sep, "/")
                info = zipfile.ZipInfo(archive_name, (1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3
                info.external_attr = 0o100644 << 16
                with open(full_path, "rb") as source, archive.open(info, "w") as destination:
                    shutil.copyfileobj(source, destination)
    return path


def create_internal_kit(version, setup_path):
    kit_path = os.path.join(DIST, "DeepExcel-Internal-v%s.zip" % version)
    if os.path.exists(kit_path):
        os.remove(kit_path)

    readme = """DeepExcel v{version} internal test build
========================================

This package is for invited test users only. It uses the private DeepExcel
internal-test publisher certificate and is not intended for public download.

Installation
------------
1. Close Microsoft Excel and WPS Spreadsheets.
2. Double-click Install-Internal-Certificate.cmd and confirm the certificate
   installation for the current Windows user.
3. Run DeepExcel.Setup.INTERNAL.exe.
4. Reopen Excel or WPS Spreadsheets and look for the DeepExcel ribbon tab.

The certificate contains only a public key. The private signing key never
leaves the DeepExcel developer's Windows certificate store.
""".format(version=version)

    entries = [
        (setup_path, "DeepExcel.Setup.INTERNAL.exe"),
        (INTERNAL_CERTIFICATE, "DeepExcel.Internal.cer"),
        (os.path.join(SCRIPTS, "install-internal-certificate.ps1"), "install-internal-certificate.ps1"),
        (os.path.join(SCRIPTS, "Install-Internal-Certificate.cmd"), "Install-Internal-Certificate.cmd"),
    ]
    with zipfile.ZipFile(kit_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for source, archive_name in entries:
            if not os.path.isfile(source):
                fail("Internal kit input is missing: %s" % source)
            info = zipfile.ZipInfo(archive_name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            with open(source, "rb") as stream:
                archive.writestr(info, stream.read())
        info = zipfile.ZipInfo("README-INTERNAL.txt", (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        info.create_system = 3
        info.external_attr = 0o100644 << 16
        archive.writestr(info, readme.encode("utf-8"))
    return kit_path


def find_iscc():
    candidates = [
        os.environ.get("INNO_COMPILER"),
        shutil.which("iscc"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"),
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
    ]
    return next((path for path in candidates if path and os.path.isfile(path)), None)


def build_installer():
    compiler = find_iscc()
    if not compiler:
        fail("Inno Setup 6 is required. Install it or set INNO_COMPILER to ISCC.exe.")
    command = [compiler]
    if signing_is_configured():
        sign_script = os.path.join(SCRIPTS, "sign.ps1")
        command.append(
            "/Sdeepsign=powershell.exe -NoProfile -ExecutionPolicy Bypass -File $q%s$q $f"
            % sign_script
        )
    command.append(os.path.join(DEPLOY, "DeepExcel.Setup.iss"))
    subprocess.run(command, check=True)
    output = os.path.join(DIST, "DeepExcel.Setup.exe")
    if not os.path.isfile(output):
        fail("Inno Setup completed without producing %s" % output)
    return output


def main():
    parser = argparse.ArgumentParser(description="Build the DeepExcel Excel + WPS release")
    parser.add_argument("--version", required=True, help="Product version, for example 0.4.17")
    parser.add_argument("--zip-only", action="store_true", help="Skip the Inno Setup installer (development only)")
    parser.add_argument(
        "--allow-unsigned",
        action="store_true",
        help="Build an unsigned installer for local validation; never distribute it",
    )
    parser.add_argument(
        "--internal",
        action="store_true",
        help="Build a clearly labelled self-signed package for invited test users",
    )
    args = parser.parse_args()

    if args.allow_unsigned and args.internal:
        fail("--allow-unsigned and --internal cannot be used together")
    if args.internal and args.zip_only:
        fail("--internal cannot be combined with --zip-only; use the labelled internal kit")
    internal_thumbprint = None
    production_thumbprint = None
    if args.internal:
        internal_thumbprint = load_internal_cert_thumbprint()
        # An internal build must use the dedicated non-exportable test key even
        # if a production credential happens to be present in the shell.
        for variable in (
            "DEEPEXCEL_PFX",
            "DEEPEXCEL_PFX_PASS",
            "USE_AZURE_TRUSTED_SIGNING",
        ):
            os.environ.pop(variable, None)
        os.environ["DEEPEXCEL_CERT_THUMBPRINT"] = internal_thumbprint
    elif args.allow_unsigned:
        # Local validation is intentionally unsigned even when the shell has a
        # signing credential. This prevents an ambiguous production-looking file.
        for variable in (
            "DEEPEXCEL_PFX",
            "DEEPEXCEL_PFX_PASS",
            "DEEPEXCEL_CERT_THUMBPRINT",
            "USE_AZURE_TRUSTED_SIGNING",
        ):
            os.environ.pop(variable, None)
    elif os.environ.get("DEEPEXCEL_CERT_THUMBPRINT"):
        fail("DEEPEXCEL_CERT_THUMBPRINT is internal-only; use --internal")

    dll_path = os.path.join(BIN_RELEASE, "DeepExcel.AddIn.dll")
    if not os.path.isfile(dll_path):
        fail("DeepExcel.AddIn.dll is missing; run scripts/_compile_only.ps1 first")
    assembly_version = read_assembly_version(dll_path)
    if not assembly_version.startswith(args.version + ".") and assembly_version != args.version:
        fail("Requested version %s does not match assembly version %s" % (args.version, assembly_version))

    print("==> DeepExcel %s unified Excel + WPS package" % args.version)
    if not args.zip_only and not args.allow_unsigned and not args.internal:
        production_thumbprint = validate_production_signing_configuration()
    generate_reg_iss(assembly_version)

    try:
        assemble_payload(args.version)
        packaged_dll = os.path.join(PAYLOAD, "Excel", "DeepExcel.AddIn.dll")
        if signing_is_configured():
            sign_files([packaged_dll])
            verify_authenticode(
                packaged_dll,
                expected_thumbprint=(
                    internal_thumbprint if args.internal else production_thumbprint
                ),
            )
        if not args.internal:
            zip_path = create_zip(args.version)
            print("[package] ZIP: %s (%.2f MB)" % (zip_path, os.path.getsize(zip_path) / 1024 / 1024))

        if args.zip_only:
            print("[package] ZIP-only development build complete")
            return

        ensure_webview2_bootstrapper()
        setup_path = build_installer()
        sign_files([setup_path])
        if args.internal:
            internal_setup_path = os.path.join(DIST, "DeepExcel.Setup.INTERNAL.exe")
            os.replace(setup_path, internal_setup_path)
            setup_path = internal_setup_path
            verify_authenticode(setup_path, expected_thumbprint=internal_thumbprint)
            kit_path = create_internal_kit(args.version, setup_path)
            print("[package] Internal kit: %s (%.2f MB)" % (kit_path, os.path.getsize(kit_path) / 1024 / 1024))
        elif args.allow_unsigned and not signing_is_configured():
            unsigned_path = os.path.join(DIST, "DeepExcel.Setup.UNSIGNED-LOCAL.exe")
            os.replace(setup_path, unsigned_path)
            setup_path = unsigned_path
        else:
            verify_authenticode(setup_path, expected_thumbprint=production_thumbprint)
        print("[package] Installer: %s (%.2f MB)" % (setup_path, os.path.getsize(setup_path) / 1024 / 1024))
        if args.internal:
            print("[package] SELF-SIGNED INTERNAL TEST BUILD - NOT FOR PUBLIC DISTRIBUTION")
        elif args.allow_unsigned and not signing_is_configured():
            print("[package] UNSIGNED LOCAL VALIDATION BUILD - DO NOT DISTRIBUTE")
        else:
            print("[package] Signed installer ready for distribution")
    finally:
        # Inno always writes this temporary production-looking name first. An
        # interrupted internal build must never leave a self-signed generic EXE.
        generic_setup = os.path.join(DIST, "DeepExcel.Setup.exe")
        if args.internal and os.path.isfile(generic_setup):
            os.remove(generic_setup)
        if os.path.isdir(PAYLOAD):
            shutil.rmtree(PAYLOAD)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        sys.exit(1)
