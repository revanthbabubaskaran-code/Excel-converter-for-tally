"""Device-bound, owner-issued licenses for Excel for Tally."""

from __future__ import annotations

import base64
import hashlib
import json
import os
from pathlib import Path
import re
import winreg

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey


PRODUCT_ID = "excel-for-tally-v1"
REQUEST_PREFIX = "EFT1-"
TOKEN_PREFIX = "EFT1"
LICENSE_FILE = "license.json"


class LicenseError(Exception):
    """Raised when a device license is invalid or cannot be accessed."""


def app_data_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base / "ExcelForTally"


def _machine_guid() -> str:
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
    except OSError as exc:
        raise LicenseError("Could not read this computer's Windows device ID.") from exc
    if not value:
        raise LicenseError("This computer does not have a Windows device ID.")
    return str(value).strip().casefold()


def device_id() -> str:
    """Return a stable, non-secret identifier unique to this Windows installation."""
    digest = hashlib.sha256(_machine_guid().encode("utf-8")).hexdigest()[:32].upper()
    return digest


def request_code() -> str:
    return f"{REQUEST_PREFIX}{device_id()}"


def _decode_urlsafe(value: str) -> bytes:
    padded = value + ("=" * (-len(value) % 4))
    return base64.urlsafe_b64decode(padded.encode("ascii"))


def _public_key() -> Ed25519PublicKey:
    public_key_file = Path(__file__).with_name("license_public_key.txt")
    try:
        encoded = public_key_file.read_text(encoding="ascii").strip()
        raw = base64.b64decode(encoded, validate=True)
        return Ed25519PublicKey.from_public_bytes(raw)
    except (OSError, ValueError) as exc:
        raise LicenseError("This app build has not been configured with its license verification key.") from exc


def verify_license(token: str, expected_device: str | None = None) -> bool:
    parts = token.strip().split(".")
    if len(parts) != 3 or parts[0] != TOKEN_PREFIX:
        return False
    try:
        payload_bytes = _decode_urlsafe(parts[1])
        signature = _decode_urlsafe(parts[2])
        _public_key().verify(signature, payload_bytes)
        payload = json.loads(payload_bytes.decode("utf-8"))
        return (
            payload.get("product") == PRODUCT_ID
            and payload.get("device_id") == (expected_device or device_id())
            and isinstance(payload.get("license_id"), str)
        )
    except (InvalidSignature, UnicodeDecodeError, json.JSONDecodeError, ValueError, LicenseError):
        return False


def saved_license() -> str | None:
    path = app_data_dir() / LICENSE_FILE
    try:
        token = path.read_text(encoding="ascii").strip()
    except FileNotFoundError:
        return None
    except OSError as exc:
        raise LicenseError("The saved license could not be read.") from exc
    return token if verify_license(token) else None


def activate_license(token: str) -> None:
    if not verify_license(token):
        raise LicenseError("This license is invalid or was issued for a different device.")
    path = app_data_dir() / LICENSE_FILE
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(token.strip(), encoding="ascii")
    except OSError as exc:
        raise LicenseError("Could not save the license on this Windows account.") from exc


def is_licensed() -> bool:
    return saved_license() is not None


def parse_request_code(value: str) -> str:
    normalized = value.strip().upper()
    if not re.fullmatch(r"EFT1-[0-9A-F]{32}", normalized):
        raise LicenseError("The license request code is invalid.")
    return normalized.removeprefix(REQUEST_PREFIX)
