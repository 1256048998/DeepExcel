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
import datetime as dt
import hashlib
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
CHECKSUM_FILE = "SHA256SUMS.txt"
UPDATE_MANIFEST_FILE = "update.json"

EXCEL_ITEMS = [
    "DeepExcel.AddIn.dll",
    "DeepExcel.AddIn.dll.config",
    # Native diagnose/repair tooling. The installer calls these instead of
    # powershell.exe, and the user gets DeepExcel.Repair.exe as the one thing to
    # run when the ribbon tab does not appear.
    "DeepExcel.Repair.exe",
    "DeepExcel.Probe32.exe",
    # Installs a staged, signature-verified update once Excel exits. Copied out
    # of here into the staging directory at update time, because Inno Setup
    # cannot overwrite a running executable.
    "DeepExcel.Updater.exe",
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
    "DeepExcel.Repair.exe",
    "DeepExcel.Probe32.exe",
    "DeepExcel.Updater.exe",
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
    # 加载项清单：声明入口 url 和按钮的 onAction 回调。开发路径的
    # build-wps.ps1 一直在复制它，发布路径却漏了，于是用户装到的加载项
    # 只有 ribbon.xml 能渲染出选项卡，所有 JS 回调都是死的——选项卡在、
    # 图标不显示、点"打开面板"毫无反应。干净沙箱 + 真实 WPS 已复现。
    "jsplugins.xml",
    "taskpane.html",
    "package.json",
    "sidecar-host.js",
    "tool-dispatcher.js",
    "wps-actions.js",
    # read_range 分页 / find 汇总的纯逻辑，wps-actions.js 依赖它
    "range-paging.js",
    "jsa-executor.js",
    # 模型配置（厂商 / 模型优先级 / API Key），与 Excel 端共用 config.json + DPAPI 凭据
    "config-store.js",
    "credential-store.js",
    "model-service.js",
    "dpapi-cli.py",
    # 对话历史 + 附件
    "conversation-store.js",
    "attachment-store.js",
    "images",
]

# WPS 端的 sidecar/ 与 web/ 不是源码，而是 Excel 端同名产物的副本。以前打包直接从
# src/DeepExcel.Wps/ 下取——那两个目录被 gitignore，是 build-wps.ps1 某次运行留下的
# 旧拷贝。2026-09-25 检查时，那份 sidecar 停在 09-13：缺 tools=[] 的纵深防御 hook、
# MaxTurns、UTF-8 流修复，发版就会把 WPS 用户带回三个安全修复之前。
# 现在从刚组装好的 Excel 载荷里复制，两个宿主拿到的侧车和面板逐字节相同。
WPS_DERIVED_FROM_EXCEL = {
    "sidecar": "sidecar",
    "web": "WebViewAssets",
}

WPS_REQUIRED_FILES = [
    "main.js",
    "ribbon.xml",
    # 这条守卫本来就是为了拦"WPS 资源缺失"，连 images/panel.svg 都列了，
    # 偏偏漏了加载项清单本身，于是缺失一路溜到用户机器上。
    "jsplugins.xml",
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

Windows SmartScreen
-------------------
DeepExcel is not code-signed yet, so Windows shows "Windows protected your PC"
the first time you run the installer. Click "More info", then "Run anyway".

You can verify the download first. Compare the output of

    certutil -hashfile DeepExcel.Setup.exe SHA256

against the value published in SHA256SUMS.txt on the download page.

If the ribbon tab does not appear, run DeepExcel.Repair.exe from the install
folder; it re-registers the add-in and writes a diagnostic bundle.

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


def copy_derived_wps_items(excel_payload, wps_payload):
    """WPS 的 sidecar/、web/ 一律取自 Excel 载荷，见 WPS_DERIVED_FROM_EXCEL。"""
    for wps_name, excel_name in WPS_DERIVED_FROM_EXCEL.items():
        source = os.path.join(excel_payload, excel_name)
        if not os.path.isdir(source):
            fail("Excel payload is missing %s, needed for WPS %s/" % (excel_name, wps_name))
        destination = os.path.join(wps_payload, wps_name)
        if os.path.exists(destination):
            shutil.rmtree(destination)
        shutil.copytree(
            source,
            destination,
            ignore=shutil.ignore_patterns("__pycache__", "tests", "*.pyc", "*.pyo"),
        )


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
    return (
        bool(os.environ.get("DEEPEXCEL_PFX"))
        or bool(os.environ.get("DEEPEXCEL_CERT_THUMBPRINT"))
        or os.environ.get("USE_AZURE_TRUSTED_SIGNING", "").lower() in ("1", "true")
    )


# Shared certificate predicate. Rejects self-signed certificates ($c.Subject -eq
# $c.Issuer) and anything that does not chain to a trusted root, so the removed
# "internal test" self-signed workflow cannot be reintroduced by setting an env
# var. An honestly unsigned build is better for users than a private root CA.
_CERT_PREDICATE = (
    "$eku=$c.Extensions|Where-Object{$_.Oid.Value -eq '2.5.29.37'}|"
    "ForEach-Object{$_.EnhancedKeyUsages}|Where-Object{$_.Value -eq '1.3.6.1.5.5.7.3.3'};"
    "if(-not $c.HasPrivateKey -or -not $eku -or $c.Subject -eq $c.Issuer "
    "-or $c.NotBefore -gt (Get-Date) -or $c.NotAfter -le (Get-Date)){exit 1};"
    "$chain=[Security.Cryptography.X509Certificates.X509Chain]::new();"
    "$chain.ChainPolicy.RevocationMode='NoCheck';"
    "if(-not $chain.Build($c)){exit 1};$c.Thumbprint"
)


def _run_cert_check(command, source_description):
    try:
        thumbprint = subprocess.check_output(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=30,
        ).strip().upper()
    except subprocess.CalledProcessError as exc:
        fail(
            "%s must be a currently valid, CA-issued code-signing certificate with a "
            "private key (self-signed certificates are rejected): %s"
            % (source_description, exc.output.strip())
        )
    if len(thumbprint) != 40 or any(ch not in "0123456789ABCDEF" for ch in thumbprint):
        fail("Cannot determine the signer thumbprint for %s" % source_description)
    return thumbprint


def validate_production_signing_configuration():
    """Return the expected signer thumbprint, or None for Azure Trusted Signing."""
    pfx = os.environ.get("DEEPEXCEL_PFX")
    store_thumbprint = (os.environ.get("DEEPEXCEL_CERT_THUMBPRINT") or "").replace(" ", "").upper()
    use_azure = os.environ.get("USE_AZURE_TRUSTED_SIGNING", "").lower() in ("1", "true")

    sources = [bool(pfx), bool(store_thumbprint), use_azure]
    if sum(1 for source in sources if source) != 1:
        fail(
            "Signed builds require exactly one signing source: DEEPEXCEL_PFX, "
            "DEEPEXCEL_CERT_THUMBPRINT, or USE_AZURE_TRUSTED_SIGNING"
        )

    if use_azure:
        # Inno Setup's preprocessor is case-sensitive; normalize values such as
        # TRUE so it also signs the nested uninstaller.
        os.environ["USE_AZURE_TRUSTED_SIGNING"] = "1"
        return None

    if store_thumbprint:
        # Certificate-store path, e.g. an EV token whose private key cannot be
        # exported to a PFX.
        if len(store_thumbprint) != 40 or any(ch not in "0123456789ABCDEF" for ch in store_thumbprint):
            fail("DEEPEXCEL_CERT_THUMBPRINT must be a 40-character SHA-1 thumbprint")
        command = (
            "$c=Get-ChildItem Cert:\\CurrentUser\\My,Cert:\\LocalMachine\\My -ErrorAction "
            "SilentlyContinue|Where-Object{($_.Thumbprint -replace '\\s','').ToUpperInvariant() "
            "-eq '%s'}|Select-Object -First 1;if(-not $c){exit 1};%s"
            % (store_thumbprint, _CERT_PREDICATE)
        )
        return _run_cert_check(command, "DEEPEXCEL_CERT_THUMBPRINT")

    if not os.path.isfile(pfx):
        fail("Production PFX does not exist: %s" % pfx)
    command = (
        "$c=[Security.Cryptography.X509Certificates.X509Certificate2]::new("
        "$env:DEEPEXCEL_PFX,$env:DEEPEXCEL_PFX_PASS,"
        "[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet);"
        + _CERT_PREDICATE
    )
    return _run_cert_check(command, "DEEPEXCEL_PFX")


def sha256_of_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_checksums(paths):
    """Publish SHA-256 digests for every distributable artifact.

    DeepExcel ships unsigned, so Authenticode cannot vouch for integrity. The
    digest file is the replacement: it is published next to the download and
    lets a user run `certutil -hashfile <file> SHA256` before installing.
    """
    existing = [path for path in paths if os.path.isfile(path)]
    if not existing:
        fail("No artifacts to checksum")
    checksum_path = os.path.join(DIST, CHECKSUM_FILE)
    lines = []
    for path in sorted(existing, key=os.path.basename):
        digest = sha256_of_file(path)
        lines.append("%s  %s" % (digest, os.path.basename(path)))
        print("[checksum] %s  %s" % (digest, os.path.basename(path)))
    # sha256sum-compatible format, LF endings so the file is stable across OSes.
    with open(checksum_path, "w", encoding="ascii", newline="\n") as stream:
        stream.write("\n".join(lines) + "\n")
    return checksum_path


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


def verify_probe_is_32bit(path):
    """A 64-bit Probe32 would test the wrong registry view and report a false pass.

    That failure is invisible at build time and only shows up as "installs fine,
    add-in missing" on a 32-bit Office machine, which is exactly the class of bug
    this tooling exists to prevent. Read the PE header directly.
    """
    with open(path, "rb") as stream:
        data = stream.read()

    def u16(offset):
        return int.from_bytes(data[offset:offset + 2], "little")

    def u32(offset):
        return int.from_bytes(data[offset:offset + 4], "little")

    if data[:2] != b"MZ":
        fail("DeepExcel.Probe32.exe is not a Windows executable")
    pe = u32(0x3C)
    if data[pe:pe + 4] != b"PE\0\0":
        fail("DeepExcel.Probe32.exe has no PE header")

    # The COFF machine field is IMAGE_FILE_MACHINE_I386 for BOTH AnyCPU and x86
    # managed assemblies, so checking it proves nothing. The real discriminator
    # is COMIMAGE_FLAGS_32BITREQUIRED in the CLI header.
    optional = pe + 24
    size_of_optional = u16(pe + 20)
    magic = u16(optional)
    if magic == 0x20B:  # PE32+
        fail("DeepExcel.Probe32.exe is a 64-bit image; rebuild with scripts/build-repair.ps1")
    if magic != 0x10B:
        fail("DeepExcel.Probe32.exe has an unrecognized optional header magic 0x%X" % magic)

    clr_directory = optional + 96 + 14 * 8
    clr_rva = u32(clr_directory)
    if clr_rva == 0:
        fail("DeepExcel.Probe32.exe has no CLI header (not a managed assembly)")

    # Translate the CLI header RVA into a file offset via the section table.
    sections = optional + size_of_optional
    clr_offset = None
    for index in range(u16(pe + 6)):
        header = sections + index * 40
        virtual_address = u32(header + 12)
        virtual_size = u32(header + 8)
        raw_pointer = u32(header + 20)
        if virtual_address <= clr_rva < virtual_address + max(virtual_size, u32(header + 16)):
            clr_offset = raw_pointer + (clr_rva - virtual_address)
            break
    if clr_offset is None:
        fail("Cannot locate the CLI header of DeepExcel.Probe32.exe")

    flags = u32(clr_offset + 16)
    if not flags & 0x00000002:  # COMIMAGE_FLAGS_32BITREQUIRED
        fail(
            "DeepExcel.Probe32.exe is AnyCPU and would run 64-bit, silently "
            "testing the wrong registry view. Rebuild with scripts/build-repair.ps1."
        )


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
    copy_derived_wps_items(excel_payload, wps_payload)

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

    # register-user.ps1 / diagnose.ps1 are no longer shipped. DeepExcel.Repair.exe
    # supersedes both for end users, and keeping PowerShell scripts out of the
    # install directory removes a standing antivirus false-positive surface.
    # They remain in scripts/ as development tools.
    verify_probe_is_32bit(os.path.join(excel_payload, "DeepExcel.Probe32.exe"))

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


def write_update_manifest(version, setup_path):
    """Publishes dist/update.json so installed clients can upgrade themselves.

    The policy is driven by what is compiled into the client, so the two can
    never be inconsistent:

      * no public key in the client -> the client cannot accept updates at all,
        so skip the manifest and say so;
      * public key present but no signing key here -> hard failure. This is the
        dangerous combination: clients are waiting for a manifest and this build
        would ship none, which looks exactly like "nobody has upgraded yet";
      * both present -> sign, then verify the result against the client's own
        public key before publishing.

    That last check is the point of the whole function. A signing key that no
    longer matches the compiled-in public key breaks updating for every user at
    once, and nothing would report it -- clients would just quietly stop
    upgrading. Catching it here costs milliseconds.
    """
    sys.path.insert(0, SCRIPTS)
    import update_signing

    embedded = update_signing.read_embedded_key()
    private_key = os.environ.get("DEEPEXCEL_UPDATE_KEY")

    if embedded is None:
        print(
            "[update] No update signing key is compiled into the client, so this\n"
            "         build cannot self-update. Generate one with:\n"
            "             python scripts/update_signing.py genkey --out <offline path>"
        )
        return None

    if not private_key:
        fail(
            "This build has an update signing public key compiled in, so its clients\n"
            "expect a signed update manifest, but DEEPEXCEL_UPDATE_KEY is not set.\n"
            "Releasing without a manifest would silently stop every client from\n"
            "upgrading. Set DEEPEXCEL_UPDATE_KEY to the private key, or pass\n"
            "--no-update-manifest to publish this build deliberately without one."
        )

    base_url = os.environ.get("DEEPEXCEL_UPDATE_BASE_URL", "").rstrip("/")
    if not base_url:
        fail(
            "DEEPEXCEL_UPDATE_BASE_URL is required to publish an update manifest.\n"
            "It is the directory the installer will be downloaded from, for example\n"
            "https://github.com/<owner>/<repo>/releases/download/v%s" % version
        )
    if not base_url.startswith("https://"):
        fail("DEEPEXCEL_UPDATE_BASE_URL must be https: %s" % base_url)

    signing_key = update_signing.describe_key(private_key)
    if signing_key["key_id"] != embedded["key_id"]:
        fail(
            "The signing key does not match the public key compiled into this build "
            "(signing %s, client expects %s). Every client would reject this release "
            "as UnknownKey. Rebuild the client after running update_signing.py genkey, "
            "or sign with the matching key."
            % (signing_key["key_id"], embedded["key_id"])
        )

    payload = update_signing.build_payload(
        version=version,
        url="%s/%s" % (base_url, os.path.basename(setup_path)),
        sha256=sha256_of_file(setup_path),
        size=os.path.getsize(setup_path),
        channel=os.environ.get("DEEPEXCEL_UPDATE_CHANNEL", "stable"),
        released_at=dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        notes=os.environ.get("DEEPEXCEL_UPDATE_NOTES", ""),
        minimum_upgradable_version=os.environ.get("DEEPEXCEL_UPDATE_MINIMUM") or None,
    )
    manifest = update_signing.sign_payload(private_key, payload)

    # Prove the published artifact verifies under the client's key before it
    # leaves this machine.
    update_signing.verify_manifest(manifest, embedded["modulus_b64"], embedded["exponent_b64"])

    path = os.path.join(DIST, UPDATE_MANIFEST_FILE)
    with open(path, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")
    print("[update] Manifest: %s (key %s, channel %s)"
          % (path, manifest["key_id"], payload["channel"]))
    print("[update] Publish it at the URL the client's UpdateSettings.FeedUrl points to,")
    print("[update] and put DeepExcel.Setup.exe at %s" % payload["url"])
    return path


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
        "--force-unsigned",
        action="store_true",
        help="Ignore any signing credential in the shell and build an unsigned installer",
    )
    parser.add_argument(
        "--no-update-manifest",
        action="store_true",
        help="Deliberately publish without dist/update.json; installed clients will not "
             "see this release",
    )
    args = parser.parse_args()

    # DeepExcel currently ships unsigned (no code-signing certificate purchased).
    # Integrity is published as SHA-256 in dist/SHA256SUMS.txt instead. Signing
    # stays fully wired so a future certificate is a pure environment change.
    production_thumbprint = None
    if args.force_unsigned:
        for variable in (
            "DEEPEXCEL_PFX",
            "DEEPEXCEL_PFX_PASS",
            "DEEPEXCEL_CERT_THUMBPRINT",
            "USE_AZURE_TRUSTED_SIGNING",
        ):
            os.environ.pop(variable, None)

    dll_path = os.path.join(BIN_RELEASE, "DeepExcel.AddIn.dll")
    if not os.path.isfile(dll_path):
        fail("DeepExcel.AddIn.dll is missing; run scripts/_compile_only.ps1 first")
    assembly_version = read_assembly_version(dll_path)
    if not assembly_version.startswith(args.version + ".") and assembly_version != args.version:
        fail("Requested version %s does not match assembly version %s" % (args.version, assembly_version))

    signed = signing_is_configured()
    print("==> DeepExcel %s unified Excel + WPS package" % args.version)
    print("[package] Signing: %s" % ("configured" if signed else "NOT configured (unsigned build)"))
    if signed and not args.zip_only:
        production_thumbprint = validate_production_signing_configuration()
    generate_reg_iss(assembly_version)

    try:
        assemble_payload(args.version)
        packaged_dll = os.path.join(PAYLOAD, "Excel", "DeepExcel.AddIn.dll")
        if signed:
            sign_files([packaged_dll])
            verify_authenticode(packaged_dll, expected_thumbprint=production_thumbprint)

        zip_path = create_zip(args.version)
        print("[package] ZIP: %s (%.2f MB)" % (zip_path, os.path.getsize(zip_path) / 1024 / 1024))

        if args.zip_only:
            write_checksums([zip_path])
            print("[package] ZIP-only development build complete")
            return

        ensure_webview2_bootstrapper()
        setup_path = build_installer()
        if signed:
            sign_files([setup_path])
            verify_authenticode(setup_path, expected_thumbprint=production_thumbprint)
        print("[package] Installer: %s (%.2f MB)" % (setup_path, os.path.getsize(setup_path) / 1024 / 1024))

        checksum_path = write_checksums([setup_path, zip_path])
        print("[package] Checksums: %s" % checksum_path)

        if args.no_update_manifest:
            print(
                "[update] Skipped by --no-update-manifest. Installed clients will not\n"
                "         see this release."
            )
        else:
            write_update_manifest(args.version, setup_path)

        if signed:
            print("[package] Signed installer ready for distribution")
        else:
            print(
                "[package] UNSIGNED build. This is the current intended release form.\n"
                "          Publish %s next to the download and tell users that\n"
                "          SmartScreen will show 'Windows protected your PC' ->\n"
                "          'More info' -> 'Run anyway'." % CHECKSUM_FILE
            )
    finally:
        if os.path.isdir(PAYLOAD):
            shutil.rmtree(PAYLOAD)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        sys.exit(1)
