"""Send short lived email verification codes through Gmail or configured SMTP."""

from __future__ import annotations

from email.message import EmailMessage
import os
import secrets
import smtplib
import ssl

from tally_final_upload_build.profile import ProfileError, normalize_email

DEFAULT_SMTP_HOST = "smtp.gmail.com"
DEFAULT_SMTP_PORT = "587"
DEFAULT_SMTP_USERNAME = "revtechsolutions.tn@gmail.com"


def send_verification_code(email: str, code: str, smtp_password: str = "") -> None:
    """Send a verification code using a supplied or environment-provided SMTP password."""
    host = os.environ.get("EXCEL_TALLY_SMTP_HOST", DEFAULT_SMTP_HOST).strip()
    port_text = os.environ.get("EXCEL_TALLY_SMTP_PORT", DEFAULT_SMTP_PORT).strip()
    username = os.environ.get("EXCEL_TALLY_SMTP_USERNAME", DEFAULT_SMTP_USERNAME).strip()
    password = (smtp_password or os.environ.get("EXCEL_TALLY_SMTP_PASSWORD", "")).replace(" ", "")
    sender = os.environ.get("EXCEL_TALLY_SMTP_FROM", username).strip()
    if not all((host, username, password, sender)):
        raise ProfileError(
            f"Enter the Google App Password for {DEFAULT_SMTP_USERNAME} to send the verification code."
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
