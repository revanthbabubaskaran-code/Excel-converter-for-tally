"""Owner tool: initialize and securely export the online service signing key."""

from __future__ import annotations

import base64
from pathlib import Path
import secrets
import sys

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src"
if str(SOURCE) not in sys.path:
    sys.path.insert(0, str(SOURCE))

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

PRIVATE_KEY_FILE = ROOT / ".license_signing_key.dpapi"
ADMIN_TOKEN_FILE = ROOT / ".license_admin_token.dpapi"
PUBLIC_KEY_FILE = ROOT / "src" / "tally_final_upload_build" / "license_public_key.txt"


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


def export_signing_key() -> None:
    from tally_final_upload_build.profile import unprotect_for_current_user

    if not PRIVATE_KEY_FILE.exists():
        raise SystemExit("Signing key not found. Run this tool with 'init' first on the owner's computer.")
    private_bytes = unprotect_for_current_user(PRIVATE_KEY_FILE.read_bytes())
    print(base64.b64encode(private_bytes).decode("ascii"))


def initialize_admin_token() -> None:
    if ADMIN_TOKEN_FILE.exists():
        raise SystemExit("An owner access token already exists; refusing to replace it.")
    from tally_final_upload_build.profile import protect_for_current_user

    token = secrets.token_urlsafe(32)
    ADMIN_TOKEN_FILE.write_bytes(protect_for_current_user(token.encode("ascii")))
    print(f"Owner access token saved encrypted for this Windows account: {ADMIN_TOKEN_FILE}")
    print("The token was not printed. Export it only when setting the Cloudflare LICENSE_ADMIN_TOKEN secret.")


def export_admin_token() -> None:
    from tally_final_upload_build.profile import unprotect_for_current_user

    if not ADMIN_TOKEN_FILE.exists():
        raise SystemExit("Owner access token not found. Run this tool with 'init-admin-token' first.")
    token = unprotect_for_current_user(ADMIN_TOKEN_FILE.read_bytes())
    print(token.decode("ascii"))


def main() -> None:
    if len(sys.argv) == 2 and sys.argv[1].casefold() == "init":
        initialize()
    elif len(sys.argv) == 2 and sys.argv[1].casefold() == "export-signing-key":
        export_signing_key()
    elif len(sys.argv) == 2 and sys.argv[1].casefold() == "init-admin-token":
        initialize_admin_token()
    elif len(sys.argv) == 2 and sys.argv[1].casefold() == "export-admin-token":
        export_admin_token()
    else:
        print("Owner online license service key tool")
        print("  uv run python tools/license_service_key.py init")
        print("  uv run python tools/license_service_key.py export-signing-key")
        print("  uv run python tools/license_service_key.py init-admin-token")
        print("  uv run python tools/license_service_key.py export-admin-token")
        raise SystemExit(2)


if __name__ == "__main__":
    main()
