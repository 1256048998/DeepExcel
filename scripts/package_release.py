#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DeepExcel release packaging script.
Usage: python scripts\package_release.py --version 0.2.0
Packages bin\Release into dist\DeepExcel-v{version}.zip
"""
import argparse
import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN_RELEASE = os.path.join(ROOT, "src", "DeepExcel.AddIn", "bin", "Release")
DIST = os.path.join(ROOT, "dist")
SCRIPTS = os.path.join(ROOT, "scripts")

# Files/dirs to include in the package
INCLUDE_ITEMS = [
    "DeepExcel.AddIn.dll",
    "DeepExcel.AddIn.dll.config",
    "Microsoft.Bcl.AsyncInterfaces.dll",
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll",
    "System.Buffers.dll",
    "System.Memory.dll",
    "System.Numerics.Vectors.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Text.Encodings.Web.dll",
    "System.Text.Json.dll",
    "System.Threading.Tasks.Extensions.dll",
    "System.ValueTuple.dll",
    "WebView2Loader.dll",
    "WebViewAssets",   # directory
    "sidecar",         # directory
]

INSTALL_TXT_TEMPLATE = """DeepExcel v{version} Installation Guide
======================================

System Requirements
-------------------
- Windows 10/11 (x64)
- Excel 2016 or later (Office 365)
- .NET Framework 4.8
- WebView2 Runtime (built-in on Win11, separate install on Win10)
- Python 3.8+ (for AI Sidecar, must be in PATH)

Installation Steps
------------------
1. Extract this ZIP to any folder (e.g. C:\\DeepExcel)

2. Open PowerShell and navigate to the extracted folder

3. Run the registration script (no admin required):
   .\\register-user.ps1

4. Launch Excel and find the "DeepExcel" tab in the ribbon

5. Click "Open Panel" button to start the AI assistant panel

6. Click the "Model" button at the top of the panel to configure API Key

Uninstallation
--------------
1. Close Excel
2. Run: .\\register-user.ps1 -Unregister
   Or manually delete registry keys:
   HKCU\\Software\\Classes\\CLSID\\{{A1B2C3D4-E5F6-4F4B-9A5F-9B3C1D2E3F4A}}
   HKCU\\Software\\Classes\\CLSID\\{{B2C3D4E5-F6A7-5B7C-AC4D-2E3F4A5B6C7D}}
   HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\DeepExcel.AddIn

Configuration Locations
-----------------------
- Config: %APPDATA%\\DeepExcel\\config.json
- API Key (encrypted): %APPDATA%\\DeepExcel\\credentials\\key_*.crypt
- Logs: %APPDATA%\\DeepExcel\\logs\\
- Snapshots: %APPDATA%\\DeepExcel\\snapshots\\

Troubleshooting
---------------
1. Add-in not visible: File -> Options -> Add-ins -> Manage: COM Add-ins -> Go -> Check DeepExcel.AddIn
2. WebView2 blank: Install WebView2 Runtime
3. Model config entry: Click "Model" button at the top of the panel

Copyright (C) 2026 DeepExcel
"""


def main():
    parser = argparse.ArgumentParser(description="Package DeepExcel release ZIP")
    parser.add_argument("--version", default="0.2.0", help="Version string (e.g. 0.2.0)")
    args = parser.parse_args()
    version = args.version

    print(f"==> DeepExcel v{version} packaging")

    if not os.path.isdir(BIN_RELEASE):
        print(f"ERROR: {BIN_RELEASE} not found. Run _compile_only.ps1 first.")
        sys.exit(1)

    os.makedirs(DIST, exist_ok=True)
    zip_path = os.path.join(DIST, f"DeepExcel-v{version}.zip")

    # Staging directory
    staging = os.path.join(os.environ.get("TEMP", "/tmp"), f"DeepExcel-pack-staging-{version}")
    if os.path.exists(staging):
        shutil.rmtree(staging)
    os.makedirs(staging)

    print("[1/3] Copying files to staging...")
    for item in INCLUDE_ITEMS:
        src = os.path.join(BIN_RELEASE, item)
        if os.path.exists(src):
            if os.path.isdir(src):
                # Filter out unwanted subdirs (tests, __pycache__) for sidecar
                ignore = shutil.ignore_patterns("__pycache__", "tests", "*.pyc")
                shutil.copytree(src, os.path.join(staging, item), ignore=ignore)
            else:
                shutil.copy2(src, staging)
        else:
            print(f"  WARNING: missing {item}")

    # Copy register-user.ps1
    reg_script_src = os.path.join(SCRIPTS, "register-user.ps1")
    if os.path.exists(reg_script_src):
        shutil.copy2(reg_script_src, staging)
    else:
        print(f"  WARNING: register-user.ps1 not found at {reg_script_src}")

    # Write INSTALL.txt
    install_path = os.path.join(staging, "INSTALL.txt")
    with open(install_path, "w", encoding="utf-8") as f:
        f.write(INSTALL_TXT_TEMPLATE.format(version=version))

    print("[2/3] Creating ZIP...")
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(staging):
            for fname in files:
                full = os.path.join(root, fname)
                rel = os.path.relpath(full, staging)
                zf.write(full, rel)

    print("[3/3] Cleaning up staging...")
    shutil.rmtree(staging)

    size_mb = os.path.getsize(zip_path) / (1024 * 1024)
    print()
    print("==> Packaging complete!")
    print(f"    File: {zip_path}")
    print(f"    Size: {size_mb:.2f} MB")
    print()
    print("Distribution: send this ZIP to users. After extraction, run register-user.ps1")


if __name__ == "__main__":
    main()
