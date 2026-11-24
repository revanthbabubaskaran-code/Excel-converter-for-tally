"""Owner tool: initialize signing keys and issue device licenses."""

from __future__ import annotations

import base64
from datetime import datetime, timezone
import json
from pathlib import Path
import secrets
import sys

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

ROOT = Path(__file__).resolve().parents[1]
PRIVATE_KEY_FILE = ROOT / ".license_signing_key.dpapi"
PUBLIC_KEY_FILE = ROOT / "src" / "tally_final_upload_build" / "license_public_key.txt"
PRODUCT_ID = "excel-for-tally-v1"
REQUEST_PREFIX = "EFT1-"


def _b64url(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def initialize() -> None:
    if PRIVATE_KEY_FILE.exists() or PUBLIC_KEY_FILE.exists():
        raise SystemExit("Signing keys already exist; refusing to replace them because it invalidates existing licenses.")
    from tally_final_upload_build.profile import protect_for_current_user

    private_key = Ed25519PrivateKey.generate()
    private_bytes = private_key.private_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PrivateFormat.Raw,
        encryption_algorithm=serialization.NoEncryption(),
    )
    public_bytes = private_key.public_key().public_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PublicFormat.Raw,
    )
    PRIVATE_KEY_FILE.write_bytes(protect_for_current_user(private_bytes))
    PUBLIC_KEY_FILE.write_text(base64.b64encode(public_bytes).decode("ascii") + "\n", encoding="ascii")
    print(f"Signing key saved on this Windows account: {PRIVATE_KEY_FILE}")
    print(f"Public verification key written to: {PUBLIC_KEY_FILE}")
    print("Keep the encrypted private key on this computer. Commit only the public key file.")


def issue(request: str) -> None:
    from tally_final_upload_build.licensing import parse_request_code
    from tally_final_upload_build.profile import unprotect_for_current_user

    device = parse_request_code(request)
    if not PRIVATE_KEY_FILE.exists():
        raise SystemExit("Signing key not found. Run this tool with 'init' first on the owner's computer.")
    private_bytes = unprotect_for_current_user(PRIVATE_KEY_FILE.read_bytes())
    private_key = Ed25519PrivateKey.from_private_bytes(private_bytes)
    payload = json.dumps(
        {
            "product": PRODUCT_ID,
            "device_id": device,
            "license_id": secrets.token_hex(8).upper(),
            "issued_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        },
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    token = f"EFT1.{_b64url(payload)}.{_b64url(private_key.sign(payload))}"
    print("\nSend this device license to the requester:\n")
    print(token)


def main() -> None:
    if len(sys.argv) == 2 and sys.argv[1].casefold() == "init":
        initialize()
    elif len(sys.argv) == 3 and sys.argv[1].casefold() == "issue":
        issue(sys.argv[2])
    else:
        print("Owner license tool")
        print("  uv run python tools/issue_license.py init")
        print("  uv run python tools/issue_license.py issue EFT1-<device-request-code>")
        raise SystemExit(2)


if __name__ == "__main__":
    main()
