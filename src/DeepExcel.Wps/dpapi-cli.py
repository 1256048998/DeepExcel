"""DeepExcel WPS 凭据读写 CLI（DPAPI，CurrentUser 作用域）。

为什么需要它：Excel 端（C# SecurityManager）把 API Key 用 DPAPI 加密后以 base64
存到 %APPDATA%/DeepExcel/credentials/key_<provider>.crypt。WPS 端是 Node 环境，
没有 DPAPI 绑定，所以借用加载项本来就依赖的 Python（ctypes 调 crypt32.dll）来做
加解密——格式与 C# 完全一致，因此两个宿主共用同一份凭据，用户只需配置一次。

用法：
    python dpapi-cli.py has    <provider>      -> 打印 1/0
    python dpapi-cli.py get    <provider>      -> 打印明文（失败打印空）
    python dpapi-cli.py set    <provider>      -> 从 stdin 读明文并加密写入
    python dpapi-cli.py delete <provider>      -> 删除凭据文件

退出码 0 表示成功，非 0 表示失败（stderr 有原因）。
"""

import base64
import ctypes
import ctypes.wintypes as wintypes
import os
import re
import sys

CRYPTPROTECT_UI_FORBIDDEN = 0x01  # 与 .NET ProtectedData 的默认行为一致（不弹 UI）
PROVIDER_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


class DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _to_blob(data):
    """构造 DATA_BLOB，同时返回底层 buffer 以保持其生命周期。"""
    buffer = ctypes.create_string_buffer(data, len(data))
    blob = DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
    return blob, buffer


def _from_blob(blob):
    try:
        return ctypes.string_at(blob.pbData, blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob.pbData)


def protect(plain_bytes):
    blob_in, _buffer = _to_blob(plain_bytes)
    blob_out = DataBlob()
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(blob_in), None, None, None, None,
        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out),
    )
    if not ok:
        raise OSError("CryptProtectData failed: %s" % ctypes.GetLastError())
    return _from_blob(blob_out)


def unprotect(cipher_bytes):
    blob_in, _buffer = _to_blob(cipher_bytes)
    blob_out = DataBlob()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None,
        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out),
    )
    if not ok:
        raise OSError("CryptUnprotectData failed: %s" % ctypes.GetLastError())
    return _from_blob(blob_out)


def credential_path(provider):
    if not PROVIDER_PATTERN.match(provider or ""):
        raise ValueError("invalid provider key")
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise OSError("APPDATA is not set")
    directory = os.path.join(appdata, "DeepExcel", "credentials")
    os.makedirs(directory, exist_ok=True)
    return os.path.join(directory, "key_%s.crypt" % provider)


def main(argv):
    if len(argv) < 3:
        sys.stderr.write("usage: dpapi-cli.py has|get|set|delete <provider>\n")
        return 2

    action, provider = argv[1], argv[2]
    try:
        path = credential_path(provider)
    except Exception as error:  # noqa: BLE001 - CLI 边界，统一报错
        sys.stderr.write("%s\n" % error)
        return 2

    if action == "has":
        exists = os.path.isfile(path) and os.path.getsize(path) > 0
        sys.stdout.write("1" if exists else "0")
        return 0

    if action == "delete":
        try:
            if os.path.isfile(path):
                os.remove(path)
            return 0
        except OSError as error:
            sys.stderr.write("%s\n" % error)
            return 1

    if action == "get":
        if not os.path.isfile(path):
            return 0  # 未配置：打印空，退出码 0
        try:
            with open(path, "r", encoding="utf-8") as handle:
                encoded = handle.read().strip()
            if not encoded:
                return 0
            # ★ fail-closed：解密失败（文件被篡改/换成明文）返回空，绝不回显原文
            plain = unprotect(base64.b64decode(encoded))
            sys.stdout.write(plain.decode("utf-8"))
            return 0
        except Exception as error:  # noqa: BLE001
            sys.stderr.write("decrypt failed: %s\n" % error)
            return 0

    if action == "set":
        plain = sys.stdin.read().strip()
        if not plain:
            sys.stderr.write("empty api key\n")
            return 2
        try:
            encoded = base64.b64encode(protect(plain.encode("utf-8"))).decode("ascii")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(encoded)
            return 0
        except Exception as error:  # noqa: BLE001
            sys.stderr.write("encrypt failed: %s\n" % error)
            return 1

    sys.stderr.write("unknown action: %s\n" % action)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
