"""Generate and send short-lived profile verification codes."""

from __future__ import annotations

import secrets

from tally_final_upload_build.gmail_oauth import GmailOAuthError, send_gmail_message
from tally_final_upload_build.profile import ProfileError, normalize_email


def send_verification_code(email: str, code: str) -> None:
    """Send a verification code through the Gmail account connected by the user."""
    try:
        recipient = normalize_email(email)
        send_gmail_message(
            to=recipient,
            subject="Your Excel for Tally verification code",
            body=(
                f"Your Excel for Tally verification code is {code}. It expires in 10 minutes. "
                "If you did not request this code, you can ignore this email."
            ),
        )
    except GmailOAuthError as exc:
        raise ProfileError(str(exc)) from exc


def new_verification_code() -> str:
    """Return a uniformly generated four digit numeric code."""
    return f"{secrets.randbelow(10_000):04d}"
