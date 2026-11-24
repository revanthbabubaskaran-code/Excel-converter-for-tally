"""Local account profile and Windows-protected remembered credentials."""

from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes
import hashlib
import hmac
import json
import os
from pathlib import Path
import re
import secrets
import tempfile
from typing import Any


PASSWORD_MIN_LENGTH = 4
PASSWORD_MAX_LENGTH = 8
PBKDF2_ITERATIONS = 310_000
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


class ProfileError(Exception):
    """Raised when a profile cannot be read, written, or validated."""


def profile_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base / "ExcelForTally" / "profile.json"


def load_profile() -> dict[str, Any] | None:
    path = profile_path()
    if not path.exists():
        return None
    try:
        profile = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(profile, dict) or not profile.get("email") or not profile.get("salt") or not profile.get("password_hash"):
            raise ValueError("Profile fields are incomplete")
        return profile
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        raise ProfileError("The saved profile could not be read. Contact support or remove the local profile to set it up again.") from exc


def save_profile(profile: dict[str, Any]) -> None:
    path = profile_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(prefix="profile-", suffix=".tmp", dir=path.parent)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(profile, handle, ensure_ascii=False, indent=2)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_name, path)
        finally:
            if os.path.exists(temp_name):
                os.unlink(temp_name)
    except OSError as exc:
        raise ProfileError(f"Could not save the profile in {path.parent}.") from exc


def normalize_email(email: str) -> str:
    normalized = email.strip().casefold()
    if len(normalized) > 254 or not EMAIL_PATTERN.fullmatch(normalized):
        raise ProfileError("Enter a valid email address.")
    return normalized


def validate_password(password: str) -> None:
    if not PASSWORD_MIN_LENGTH <= len(password) <= PASSWORD_MAX_LENGTH:
        raise ProfileError("Password must be 4 to 8 characters long.")


def create_profile(email: str, password: str) -> dict[str, Any]:
    normalized_email = normalize_email(email)
    validate_password(password)
    salt = secrets.token_bytes(16)
    password_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    profile: dict[str, Any] = {
        "email": normalized_email,
        "salt": salt.hex(),
        "password_hash": password_hash.hex(),
        "remember_me": False,
        "remembered_password": None,
    }
    save_profile(profile)
    return profile


def password_matches(profile: dict[str, Any], email: str, password: str) -> bool:
    try:
        normalized_email = normalize_email(email)
        validate_password(password)
        salt = bytes.fromhex(profile["salt"])
        expected = bytes.fromhex(profile["password_hash"])
    except (ProfileError, KeyError, TypeError, ValueError):
        return False
    actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return hmac.compare_digest(normalized_email, str(profile["email"]).casefold()) and hmac.compare_digest(actual, expected)


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob(data: bytes) -> tuple[_DataBlob, ctypes.Array]:
    buffer = ctypes.create_string_buffer(data, len(data))
    return _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def protect_for_current_user(data: bytes) -> bytes:
    if os.name != "nt":
        raise ProfileError("Remember me is supported only on Windows.")
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DataBlob), wintypes.LPCWSTR, ctypes.POINTER(_DataBlob),
        wintypes.LPVOID, wintypes.LPVOID, wintypes.DWORD, ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    source, keep_alive = _blob(data)
    result = _DataBlob()
    if not crypt32.CryptProtectData(ctypes.byref(source), "Excel for Tally", None, None, None, 0x1, ctypes.byref(result)):
        raise ProfileError("Windows could not securely save the remembered password.")
    try:
        return ctypes.string_at(result.pbData, result.cbData)
    finally:
        kernel32.LocalFree(result.pbData)


def unprotect_for_current_user(data: bytes) -> bytes:
    if os.name != "nt":
        raise ProfileError("Remember me is supported only on Windows.")
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DataBlob), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(_DataBlob),
        wintypes.LPVOID, wintypes.LPVOID, wintypes.DWORD, ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    source, keep_alive = _blob(data)
    result = _DataBlob()
    if not crypt32.CryptUnprotectData(ctypes.byref(source), None, None, None, None, 0x1, ctypes.byref(result)):
        raise ProfileError("Windows could not unlock the remembered password for this account.")
    try:
        return ctypes.string_at(result.pbData, result.cbData)
    finally:
        kernel32.LocalFree(result.pbData)


def remembered_password(profile: dict[str, Any]) -> str | None:
    encoded = profile.get("remembered_password")
    if not profile.get("remember_me") or not encoded:
        return None
    encrypted = base64.b64decode(encoded, validate=True)
    return unprotect_for_current_user(encrypted).decode("utf-8")


def set_remembered_login(profile: dict[str, Any], password: str, remember: bool) -> None:
    updated = dict(profile)
    updated["remember_me"] = bool(remember)
    if remember:
        encrypted = protect_for_current_user(password.encode("utf-8"))
        updated["remembered_password"] = base64.b64encode(encrypted).decode("ascii")
    else:
        updated["remembered_password"] = None
    save_profile(updated)
    profile.clear()
    profile.update(updated)
