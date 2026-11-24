"""Connect a Gmail account in the system browser and send through the Gmail API."""

from __future__ import annotations

import base64
from email.message import EmailMessage
import hashlib
import json
import os
from pathlib import Path
import secrets
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urlsplit
from urllib.request import Request, urlopen
import webbrowser

from tally_final_upload_build.profile import (
    ProfileError,
    profile_path,
    protect_for_current_user,
    unprotect_for_current_user,
)


GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_SEND_URL = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
OAUTH_SCOPE = "https://www.googleapis.com/auth/gmail.send"
CLIENT_ID_FILE = "google_oauth_client_id.txt"
TOKEN_FILE = "gmail_credentials.dpapi"
OAUTH_TIMEOUT_SECONDS = 300


class GmailOAuthError(Exception):
    """Raised when Gmail account connection or message delivery fails."""


def _app_data_dir() -> Path:
    return profile_path().parent


def gmail_client_id() -> str | None:
    configured = os.environ.get("EXCEL_TALLY_GOOGLE_CLIENT_ID", "").strip()
    if configured:
        return configured
    try:
        value = (_app_data_dir() / CLIENT_ID_FILE).read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return None
    except OSError as exc:
        raise GmailOAuthError("Could not read the Google OAuth client ID setting.") from exc
    return value or None


def save_gmail_client_id(client_id: str) -> None:
    value = client_id.strip()
    if not value.endswith(".apps.googleusercontent.com"):
        raise GmailOAuthError("Enter the OAuth client ID for a Google Desktop app.")
    path = _app_data_dir() / CLIENT_ID_FILE
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(value, encoding="utf-8")
    except OSError as exc:
        raise GmailOAuthError("Could not save the Google OAuth client ID.") from exc


def _token_path() -> Path:
    return _app_data_dir() / TOKEN_FILE


def _load_credentials() -> dict[str, Any] | None:
    try:
        encrypted = _token_path().read_bytes()
    except FileNotFoundError:
        return None
    except OSError as exc:
        raise GmailOAuthError("Could not read the saved Gmail connection.") from exc
    try:
        payload = json.loads(unprotect_for_current_user(encrypted).decode("utf-8"))
        if not payload.get("client_id") or not payload.get("refresh_token"):
            raise ValueError("Incomplete credentials")
        return payload
    except (ProfileError, UnicodeDecodeError, json.JSONDecodeError, AttributeError, ValueError) as exc:
        raise GmailOAuthError("The saved Gmail connection could not be unlocked. Connect the account again.") from exc


def _save_credentials(client_id: str, refresh_token: str) -> None:
    payload = json.dumps({"client_id": client_id, "refresh_token": refresh_token}).encode("utf-8")
    path = _token_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(protect_for_current_user(payload))
    except (OSError, ProfileError) as exc:
        raise GmailOAuthError("Could not securely save the Gmail connection on this Windows account.") from exc


def is_gmail_connected() -> bool:
    return _load_credentials() is not None


def _post_form(url: str, fields: dict[str, str]) -> dict[str, Any]:
    data = urlencode(fields).encode("ascii")
    request = Request(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}, method="POST")
    try:
        with urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        try:
            details = json.loads(exc.read().decode("utf-8"))
            message = details.get("error_description") or details.get("error") or "Google rejected the request."
        except (UnicodeDecodeError, json.JSONDecodeError):
            message = "Google rejected the request."
        raise GmailOAuthError(str(message)) from exc
    except (URLError, TimeoutError, OSError) as exc:
        raise GmailOAuthError("Could not connect to Google. Check the internet connection and try again.") from exc
    except json.JSONDecodeError as exc:
        raise GmailOAuthError("Google returned an unreadable response.") from exc
    if not isinstance(payload, dict):
        raise GmailOAuthError("Google returned an invalid response.")
    return payload


def authorize_gmail(client_id: str) -> None:
    """Complete OAuth authorization via default browser and a loopback callback."""
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode("ascii")).digest()).decode("ascii").rstrip("=")
    state = secrets.token_urlsafe(32)
    received: dict[str, str] = {}

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 - stdlib handler name
            params = parse_qs(urlsplit(self.path).query)
            returned_state = params.get("state", [""])[0]
            if not secrets.compare_digest(returned_state, state):
                received["error"] = "The Google authorization response did not match this request."
            elif params.get("error"):
                received["error"] = "Google sign-in was cancelled or denied."
            else:
                received["code"] = params.get("code", [""])[0]
            content = (
                "<!doctype html><html><meta charset='utf-8'><title>Excel for Tally</title>"
                "<body style='font:16px Segoe UI,sans-serif;padding:40px'>"
                "Gmail connection finished. Return to Excel for Tally.</body></html>"
            ).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)

        def log_message(self, _format: str, *_args: Any) -> None:
            return

    try:
        server = ThreadingHTTPServer(("127.0.0.1", 0), CallbackHandler)
    except OSError as exc:
        raise GmailOAuthError("Could not start the local Google sign-in callback.") from exc
    server.timeout = 1
    redirect_uri = f"http://127.0.0.1:{server.server_address[1]}"
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": OAUTH_SCOPE,
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    auth_url = f"{GOOGLE_AUTH_URL}?{urlencode(params)}"
    try:
        if not webbrowser.open(auth_url, new=1, autoraise=True):
            raise GmailOAuthError("Could not open the default browser. Open the sign-in link manually or check browser settings.")
        deadline = time.monotonic() + OAUTH_TIMEOUT_SECONDS
        while "code" not in received and "error" not in received and time.monotonic() < deadline:
            server.handle_request()
    finally:
        server.server_close()
    if received.get("error"):
        raise GmailOAuthError(received["error"])
    auth_code = received.get("code")
    if not auth_code:
        raise GmailOAuthError("Google sign-in timed out. Click Connect account and try again.")

    token_response = _post_form(
        GOOGLE_TOKEN_URL,
        {
            "client_id": client_id,
            "code": auth_code,
            "code_verifier": verifier,
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri,
        },
    )
    refresh_token = token_response.get("refresh_token")
    if not refresh_token:
        raise GmailOAuthError("Google did not return a refresh token. Reconnect and approve the requested permissions.")
    _save_credentials(client_id, str(refresh_token))


def _access_token(credentials: dict[str, Any]) -> str:
    response = _post_form(
        GOOGLE_TOKEN_URL,
        {
            "client_id": str(credentials["client_id"]),
            "refresh_token": str(credentials["refresh_token"]),
            "grant_type": "refresh_token",
        },
    )
    token = response.get("access_token")
    if not token:
        raise GmailOAuthError("Gmail authorization expired. Click Connect account and reconnect.")
    return str(token)


def send_gmail_message(to: str, subject: str, body: str) -> None:
    credentials = _load_credentials()
    if not credentials:
        raise GmailOAuthError("Connect a Gmail account before sending a verification code.")
    access_token = _access_token(credentials)
    message = EmailMessage()
    message["To"] = to
    message["Subject"] = subject
    message.set_content(body)
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii").rstrip("=")
    request = Request(
        GMAIL_SEND_URL,
        data=json.dumps({"raw": raw_message}).encode("utf-8"),
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=30) as response:
            response.read()
    except HTTPError as exc:
        if exc.code == 401:
            raise GmailOAuthError("Gmail authorization expired. Click Connect account and reconnect.") from exc
        raise GmailOAuthError("Gmail could not send the verification email. Check account access and try again.") from exc
    except (URLError, TimeoutError, OSError) as exc:
        raise GmailOAuthError("Could not connect to Gmail. Check the internet connection and try again.") from exc
