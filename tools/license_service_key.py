"""Owner tool: initialize and securely export the online service signing key."""

from __future__ import annotations

import base64
from pathlib import Path
import sys

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

ROOT = Path(__file__).resolve().parents[1]
PRIVATE_KEY_FILE = ROOT / ".license_signing_key.dpapi"
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


def main() -> None:
    if len(sys.argv) == 2 and sys.argv[1].casefold() == "init":
        initialize()
    elif len(sys.argv) == 2 and sys.argv[1].casefold() == "export-signing-key":
        export_signing_key()
    else:
        print("Owner online license service key tool")
        print("  uv run python tools/license_service_key.py init")
        print("  uv run python tools/license_service_key.py export-signing-key")
        raise SystemExit(2)


if __name__ == "__main__":
    main()
