#!/usr/bin/env python3
"""Atomically install/remove the DeepExcel WPS publish.xml entry."""

import argparse
import hashlib
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET


VERSION_RE = re.compile(r"^[0-9]+(?:\.[0-9]+){1,3}$")
PAYLOAD_RE = re.compile(r"^DeepExcel_[0-9]+(?:\.[0-9]+){1,3}$")


def file_hash(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", required=True, choices=("install", "uninstall"), type=str.lower)
    parser.add_argument("--version", required=True)
    parser.add_argument("--root", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if not VERSION_RE.fullmatch(args.version):
        parser.error("invalid version")
    return args


def main():
    args = parse_args()
    appdata = os.environ.get("APPDATA")
    if args.root:
        root = os.path.abspath(args.root)
    elif appdata:
        root = os.path.abspath(os.path.join(appdata, "kingsoft", "wps", "jsaddons"))
    else:
        raise RuntimeError("APPDATA is not available")
    os.makedirs(root, exist_ok=True)

    manifest = os.path.join(root, "publish.xml")
    original_exists = os.path.isfile(manifest)
    original_hash = file_hash(manifest) if original_exists else None

    if original_exists:
        try:
            tree = ET.parse(manifest)
        except Exception:
            invalid_backup = manifest + ".invalid.bak"
            shutil.copy2(manifest, invalid_backup)
            raise RuntimeError("existing WPS publish.xml is invalid; backup: " + invalid_backup)
        root_element = tree.getroot()
    else:
        root_element = ET.Element("jsplugins")
        tree = ET.ElementTree(root_element)

    if root_element.tag != "jsplugins":
        raise RuntimeError("unexpected WPS manifest root (expected jsplugins)")

    for child in list(root_element):
        if child.get("name") == "DeepExcel":
            root_element.remove(child)

    if args.mode == "install":
        ET.SubElement(root_element, "jsplugin", {
            "name": "DeepExcel",
            "type": "et",
            "url": "DeepExcel_" + args.version,
            "version": args.version,
            "enable": "enable_dev",
            "install": "null",
            "customDomain": "",
        })

    try:
        ET.indent(tree, space="  ")
    except AttributeError:
        pass

    fd, temporary = tempfile.mkstemp(prefix=".publish.", suffix=".tmp", dir=root)
    os.close(fd)
    try:
        tree.write(temporary, encoding="utf-8", xml_declaration=True)
        ET.parse(temporary)  # validate completed output before commit

        exists_now = os.path.isfile(manifest)
        if exists_now != original_exists:
            raise RuntimeError("WPS publish.xml changed during update; retry")
        if original_exists and file_hash(manifest) != original_hash:
            raise RuntimeError("WPS publish.xml was modified by another process; retry")

        if original_exists:
            shutil.copy2(manifest, manifest + ".bak")
        os.replace(temporary, manifest)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)

    if args.mode == "install":
        current = "DeepExcel_" + args.version
        for entry in os.scandir(root):
            if (entry.is_dir(follow_symlinks=False) and not entry.is_symlink()
                    and PAYLOAD_RE.fullmatch(entry.name) and entry.name != current):
                try:
                    shutil.rmtree(entry.path)
                except OSError as exc:
                    print("warning: old WPS payload cleanup failed: " + str(exc), file=sys.stderr)

    print("WPS publish.xml updated: %s DeepExcel %s" % (args.mode, args.version))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("WPS manifest update failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
