"""Send short lived email verification codes using configured SMTP."""

from __future__ import annotations

from email.message import EmailMessage
import os
import secrets
import smtplib
import ssl

from tally_final_upload_build.profile import ProfileError, normalize_email


def send_verification_code(email: str, code: str) -> None:
    """Send a verification code; credentials must come from environment variables."""
    host = os.environ.get("EXCEL_TALLY_SMTP_HOST", "").strip()
    port_text = os.environ.get("EXCEL_TALLY_SMTP_PORT", "587").strip()
    username = os.environ.get("EXCEL_TALLY_SMTP_USERNAME", "").strip()
    password = os.environ.get("EXCEL_TALLY_SMTP_PASSWORD", "")
    sender = os.environ.get("EXCEL_TALLY_SMTP_FROM", username).strip()
    if not all((host, username, password, sender)):
        raise ProfileError(
            "Email verification is not configured. Set EXCEL_TALLY_SMTP_HOST, "
            "EXCEL_TALLY_SMTP_PORT, EXCEL_TALLY_SMTP_USERNAME, "
            "EXCEL_TALLY_SMTP_PASSWORD, and EXCEL_TALLY_SMTP_FROM, then restart the app."
        )
    try:
        port = int(port_text)
        recipient = normalize_email(email)
    except (ValueError, ProfileError) as exc:
        raise ProfileError("Email verification settings or email address are invalid.") from exc

    message = EmailMessage()
    message["Subject"] = "Your Excel for Tally verification code"
    message["From"] = sender
    message["To"] = recipient
    message.set_content(
        f"Your Excel for Tally verification code is {code}. It expires in 10 minutes. "
        "If you did not request this code, you can ignore this email."
    )

    try:
        context = ssl.create_default_context()
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=20, context=context) as smtp:
                smtp.login(username, password)
                smtp.send_message(message)
        else:
            with smtplib.SMTP(host, port, timeout=20) as smtp:
                smtp.ehlo()
                smtp.starttls(context=context)
                smtp.ehlo()
                smtp.login(username, password)
                smtp.send_message(message)
    except (OSError, smtplib.SMTPException) as exc:
        raise ProfileError("Could not send the verification email. Check SMTP settings and try again.") from exc


def new_verification_code() -> str:
    """Return a uniformly generated four digit numeric code."""
    return f"{secrets.randbelow(10_000):04d}"
