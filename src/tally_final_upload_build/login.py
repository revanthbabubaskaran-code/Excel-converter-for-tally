"""First-run profile setup and local login window."""

from __future__ import annotations

import hashlib
import hmac
import time
import tkinter as tk
from tkinter import messagebox, ttk

from tally_final_upload_build.app import ExcelForTallyApp
from tally_final_upload_build.email_verification import new_verification_code, send_verification_code
from tally_final_upload_build.profile import (
    PASSWORD_MAX_LENGTH,
    PASSWORD_MIN_LENGTH,
    ProfileError,
    create_profile,
    load_profile,
    normalize_email,
    password_matches,
    remembered_password,
    set_remembered_login,
)


class LoginFlow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.profile: dict | None = None
        self.password_vcmd = (root.register(self._password_length_ok), "%P")
        self._configure()
        try:
            self.profile = load_profile()
        except ProfileError as exc:
            messagebox.showerror("Profile could not be loaded", f"{exc}\n\nProfile location:\n{self._profile_location()}", parent=root)
            root.destroy()
            return
        if self.profile is None:
            self.show_setup()
        else:
            self.show_login()

    @staticmethod
    def _password_length_ok(proposed: str) -> bool:
        return len(proposed) <= PASSWORD_MAX_LENGTH

    @staticmethod
    def _verification_code_ok(proposed: str) -> bool:
        return len(proposed) <= 4 and (not proposed or proposed.isascii() and proposed.isdigit())

    @staticmethod
    def _profile_location() -> str:
        from tally_final_upload_build.profile import profile_path

        return str(profile_path())

    def _configure(self) -> None:
        self.root.title("Excel for Tally — Sign in")
        self.root.geometry("520x520")
        self.root.minsize(480, 440)
        self.root.configure(bg="#f3f6f8")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f6f8")
        style.configure("TLabel", background="#f3f6f8", foreground="#23313a", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#123d54")
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#52636d")
        style.configure("Error.TLabel", foreground="#c62828", background="#f3f6f8", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9))
        style.configure("TEntry", padding=(8, 7), font=("Segoe UI", 11))

    def _clear(self) -> None:
        for child in self.root.winfo_children():
            child.destroy()

    def _content(self) -> ttk.Frame:
        self._clear()
        self.root.geometry("520x520")
        outer = ttk.Frame(self.root, padding=(36, 28, 36, 30))
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Excel for Tally", style="Title.TLabel").pack(anchor="w")
        return outer

    def _password_entry(self, parent: ttk.Frame) -> ttk.Entry:
        return ttk.Entry(
            parent,
            show="*",
            validate="key",
            validatecommand=self.password_vcmd,
        )

    def show_setup(self) -> None:
        outer = self._content()
        self.root.geometry("560x760")
        self.root.minsize(520, 700)
        ttk.Label(outer, text="Create your local profile", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(20, 3))
        ttk.Label(
            outer,
            text="Verify your email before creating a profile. Your profile and password hash are saved on this Windows account.",
            style="Sub.TLabel",
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(0, 16))

        ttk.Label(outer, text="Login email ID").pack(anchor="w")
        email_entry = ttk.Entry(outer, font=("Segoe UI", 11))
        email_entry.pack(fill="x", pady=(4, 12))

        ttk.Label(outer, text="Outgoing email: revtechsolutions.tn@gmail.com", style="Sub.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Google App Password").pack(anchor="w", pady=(8, 0))
        smtp_password_entry = ttk.Entry(outer, font=("Segoe UI", 11), show="*")
        smtp_password_entry.pack(fill="x", pady=(4, 10))
        ttk.Label(
            outer,
            text="Enter an App Password for the sending Gmail account. It is used only to send this email and is not saved in the profile.",
            style="Sub.TLabel",
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        code_frame = ttk.Frame(outer)
        ttk.Label(code_frame, text="4-digit email verification code").pack(anchor="w")
        code_vcmd = (self.root.register(self._verification_code_ok), "%P")
        code_entry = ttk.Entry(code_frame, font=("Segoe UI", 11), validate="key", validatecommand=code_vcmd)
        code_entry.pack(fill="x", pady=(4, 6))
        code_status = ttk.Label(code_frame, text="", style="Sub.TLabel", wraplength=440)
        code_status.pack(anchor="w")

        verified_email: str | None = None
        expected_code_hash: bytes | None = None
        expires_at = 0.0

        def request_code() -> None:
            nonlocal verified_email, expected_code_hash, expires_at
            verified_email = None
            expected_code_hash = None
            code_entry.delete(0, "end")
            code_status.configure(text="")
            try:
                email = normalize_email(email_entry.get())
                code = new_verification_code()
                send_verification_code(email, code, smtp_password_entry.get())
            except ProfileError as exc:
                error.configure(text=str(exc))
                return
            except Exception:
                error.configure(text="Could not send the verification email. Check your connection and settings.")
                return
            expected_code_hash = hashlib.sha256(code.encode("ascii")).digest()
            verified_email = email
            expires_at = time.monotonic() + 600
            code_status.configure(text="A verification code was sent. It expires in 10 minutes.")
            error.configure(text="")
            verify_button.configure(state="normal")
            code_entry.focus_set()

        send_button = ttk.Button(outer, text="Send verification code", command=request_code)
        send_button.pack(anchor="w", pady=(0, 8))
        code_frame.pack(fill="x", pady=(0, 10))

        def verify_code() -> None:
            nonlocal verified_email, expected_code_hash
            if not verified_email or expected_code_hash is None:
                error.configure(text="Send a verification code first.")
                return
            if time.monotonic() > expires_at:
                verified_email = None
                expected_code_hash = None
                error.configure(text="The verification code expired. Request a new code.")
                return
            try:
                entered_email = normalize_email(email_entry.get())
            except ProfileError as exc:
                error.configure(text=str(exc))
                return
            entered_hash = hashlib.sha256(code_entry.get().strip().encode("utf-8")).digest()
            if not hmac.compare_digest(entered_email, verified_email) or not hmac.compare_digest(entered_hash, expected_code_hash):
                error.configure(text="Verification code is incorrect.")
                return
            error.configure(text="Email verified. You can create your profile.", style="Sub.TLabel")
            verify_button.configure(state="disabled")

        verify_button = ttk.Button(code_frame, text="Verify code", command=verify_code)
        verify_button.pack(anchor="e", pady=(6, 0))
        verify_button.configure(state="disabled")

        ttk.Label(outer, text=f"Password ({PASSWORD_MIN_LENGTH}–{PASSWORD_MAX_LENGTH} characters)").pack(anchor="w")
        password_entry = self._password_entry(outer)
        password_entry.pack(fill="x", pady=(4, 12))

        ttk.Label(outer, text="Confirm password").pack(anchor="w")
        confirm_entry = self._password_entry(outer)
        confirm_entry.pack(fill="x", pady=(4, 7))

        error = ttk.Label(outer, text="", style="Error.TLabel", wraplength=440)
        error.pack(anchor="w", pady=(4, 4))

        def create() -> None:
            try:
                email = normalize_email(email_entry.get())
            except ProfileError as exc:
                error.configure(text=str(exc), style="Error.TLabel")
                return
            if verified_email != email or expected_code_hash is None or time.monotonic() > expires_at or code_entry.get().strip() == "":
                error.configure(text="Verify the code sent to this email before creating your profile.", style="Error.TLabel")
                return
            entered_hash = hashlib.sha256(code_entry.get().strip().encode("utf-8")).digest()
            if not hmac.compare_digest(entered_hash, expected_code_hash):
                error.configure(text="Verification code is incorrect.", style="Error.TLabel")
                return
            password = password_entry.get()
            if password != confirm_entry.get():
                error.configure(text="Passwords do not match.")
                return
            try:
                self.profile = create_profile(email, password)
            except ProfileError as exc:
                error.configure(text=str(exc))
                return
            self.show_login(info="Profile created. Sign in to continue.")

        button = ttk.Button(outer, text="Create profile", style="Accent.TButton", command=create)
        button.pack(anchor="e", pady=(13, 0))
        confirm_entry.bind("<Return>", lambda _event: create())
        email_entry.focus_set()

    def show_login(self, info: str = "") -> None:
        if self.profile is None:
            self.show_setup()
            return
        outer = self._content()
        ttk.Label(outer, text="Sign in", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(20, 3))
        ttk.Label(
            outer,
            text="Enter the email ID and password saved in your local profile.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(0, 18))

        email_entry = ttk.Entry(outer, font=("Segoe UI", 11))
        password_entry = self._password_entry(outer)
        ttk.Label(outer, text="Login email ID").pack(anchor="w")
        email_entry.pack(fill="x", pady=(4, 13))
        ttk.Label(outer, text=f"Password ({PASSWORD_MIN_LENGTH}–{PASSWORD_MAX_LENGTH} characters)").pack(anchor="w")
        password_entry.pack(fill="x", pady=(4, 6))

        remember_default = bool(self.profile.get("remember_me"))
        remember_var = tk.BooleanVar(value=remember_default)
        if remember_default:
            email_entry.insert(0, str(self.profile.get("email", "")))
            try:
                saved_password = remembered_password(self.profile)
            except (ProfileError, ValueError) as exc:
                saved_password = None
                info = "The saved password could not be unlocked. Enter it again."
            if saved_password is not None:
                password_entry.insert(0, saved_password)

        ttk.Checkbutton(outer, text="Remember me on this Windows account", variable=remember_var).pack(anchor="w", pady=(1, 12))
        error = ttk.Label(outer, text=info, style="Error.TLabel" if info.startswith("The saved") else "Sub.TLabel", wraplength=440)
        error.pack(anchor="w", pady=(2, 5))

        def login() -> None:
            email = email_entry.get()
            password = password_entry.get()
            if not password_matches(self.profile, email, password):
                error.configure(text="login Email / password is incorrect", style="Error.TLabel")
                return
            try:
                set_remembered_login(self.profile, password, remember_var.get())
            except ProfileError as exc:
                remember_var.set(False)
                try:
                    set_remembered_login(self.profile, password, False)
                except ProfileError:
                    pass
                messagebox.showwarning(
                    "Remember me unavailable",
                    f"{exc}\n\nYou can still sign in, but the password will not be saved for autofill.",
                    parent=self.root,
                )
            self.show_converter()

        button = ttk.Button(outer, text="Log in", style="Accent.TButton", command=login)
        button.pack(anchor="e", pady=(16, 0))
        password_entry.bind("<Return>", lambda _event: login())
        (password_entry if remember_default and password_entry.get() else email_entry).focus_set()

    def show_converter(self) -> None:
        self._clear()
        self.root.geometry("820x650")
        ExcelForTallyApp(self.root, on_logout=self.show_login)


def main() -> None:
    root = tk.Tk()
    LoginFlow(root)
    root.mainloop()
