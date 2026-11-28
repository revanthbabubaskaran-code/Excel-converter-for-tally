"""First-run profile setup and local login window."""

from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk

from tally_final_upload_build.app import ExcelForTallyApp
from tally_final_upload_build.licensing import (
    LicenseError,
    activate_license,
    check_online_license_request,
    is_licensed,
    license_service_url,
    request_code,
    save_license_service_url,
    submit_online_license_request,
)
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
        try:
            licensed = is_licensed()
        except LicenseError as exc:
            messagebox.showerror("License could not be loaded", str(exc), parent=root)
            root.destroy()
            return
        if not licensed:
            self.show_license()
        elif self.profile is None:
            self.show_setup()
        else:
            self.show_login()

    @staticmethod
    def _password_length_ok(proposed: str) -> bool:
        return len(proposed) <= PASSWORD_MAX_LENGTH

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
        self.root.geometry("560x560")
        self.root.minsize(520, 500)
        ttk.Label(outer, text="Create your local profile", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(20, 3))
        ttk.Label(
            outer,
            text="Create a profile with your email ID and a password. Your profile and password hash are saved on this Windows account.",
            style="Sub.TLabel",
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(0, 16))

        ttk.Label(outer, text="Login email ID").pack(anchor="w")
        email_entry = ttk.Entry(outer, font=("Segoe UI", 11))
        email_entry.pack(fill="x", pady=(4, 12))

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

    def show_license(self) -> None:
        outer = self._content()
        self.root.geometry("560x600")
        self.root.minsize(520, 540)
        ttk.Label(outer, text="License this device", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(20, 3))
        ttk.Label(
            outer,
            text="This copy needs a license for this Windows device. Send the request code below to the software owner to receive a device license.",
            style="Sub.TLabel",
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        ttk.Label(outer, text="Device license request code").pack(anchor="w")
        try:
            code = request_code()
        except LicenseError as exc:
            code = ""
            ttk.Label(outer, text=str(exc), style="Error.TLabel", wraplength=440).pack(anchor="w", pady=(5, 8))
        request_entry = ttk.Entry(outer, font=("Consolas", 10))
        request_entry.pack(fill="x", pady=(4, 7))
        if code:
            request_entry.insert(0, code)

        def copy_request() -> None:
            if not code:
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            self.root.update()
            status.configure(text="Request code copied.", style="Sub.TLabel")

        ttk.Button(outer, text="Copy request code", command=copy_request).pack(anchor="e", pady=(0, 14))
        ttk.Label(outer, text="Online license service address").pack(anchor="w")
        service_entry = ttk.Entry(outer, font=("Segoe UI", 10))
        service_entry.pack(fill="x", pady=(4, 8))
        try:
            configured_service = license_service_url()
        except LicenseError as exc:
            configured_service = None
            ttk.Label(outer, text=str(exc), style="Error.TLabel", wraplength=440).pack(anchor="w")
        if configured_service:
            service_entry.insert(0, configured_service)

        status = ttk.Label(outer, text="Requests are reviewed by the software owner.", style="Sub.TLabel", wraplength=440)
        status.pack(anchor="w", pady=(4, 8))

        def finish_service_call(result: dict | None, failure: str, checking: bool) -> None:
            request_button.configure(state="normal")
            check_button.configure(state="normal")
            if failure:
                status.configure(text=failure, style="Error.TLabel")
                return
            if result is None:
                status.configure(text="The license service returned no result.", style="Error.TLabel")
                return
            state = result.get("status")
            if state == "approved" and result.get("license_token"):
                try:
                    activate_license(str(result["license_token"]))
                except LicenseError as exc:
                    status.configure(text=str(exc), style="Error.TLabel")
                    return
                if self.profile is None:
                    self.show_setup()
                else:
                    self.show_login(info="Device license activated.")
            elif state == "pending":
                status.configure(text="Request submitted. Waiting for owner approval. Click Check status after approval.", style="Sub.TLabel")
            elif state == "rejected":
                status.configure(text="The license request was declined. You may submit it again for review.", style="Error.TLabel")
            elif state == "not_found" and checking:
                status.configure(text="No request found. Click Request license first.", style="Error.TLabel")
            else:
                status.configure(text="The online service returned an unexpected license status.", style="Error.TLabel")

        def run_service_call(checking: bool) -> None:
            try:
                service_url = save_license_service_url(service_entry.get())
                request_button.configure(state="disabled")
                check_button.configure(state="disabled")
                status.configure(text="Contacting the online license service…", style="Sub.TLabel")
            except LicenseError as exc:
                status.configure(text=str(exc), style="Error.TLabel")
                return

            def worker() -> None:
                result: dict | None = None
                failure = ""
                try:
                    if checking:
                        result = check_online_license_request(code, service_url)
                    else:
                        result = submit_online_license_request(code, service_url)
                except LicenseError as exc:
                    failure = str(exc)
                self.root.after(0, lambda: finish_service_call(result, failure, checking))

            threading.Thread(target=worker, daemon=True).start()

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(8, 0))
        check_button = ttk.Button(buttons, text="Check status", command=lambda: run_service_call(True))
        check_button.pack(side="left")
        request_button = ttk.Button(buttons, text="Request license", style="Accent.TButton", command=lambda: run_service_call(False))
        request_button.pack(side="right")

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
            self.show_profiles()

        button = ttk.Button(outer, text="Log in", style="Accent.TButton", command=login)
        button.pack(anchor="e", pady=(16, 0))
        password_entry.bind("<Return>", lambda _event: login())
        (password_entry if remember_default and password_entry.get() else email_entry).focus_set()

    def show_profiles(self) -> None:
        outer = self._content()
        self.root.title("Excel for Tally — Choose input type")
        self.root.geometry("560x480")
        self.root.minsize(520, 440)
        ttk.Label(outer, text="Choose input type", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", pady=(20, 4)
        )
        ttk.Label(
            outer,
            text="Select the type of reports you want to convert.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(0, 20))

        profiles = (("Hospital", self.show_converter),)
        for label, open_profile in profiles:
            ttk.Button(
                outer,
                text=label,
                style="Accent.TButton",
                command=open_profile,
            ).pack(fill="x", pady=(0, 10))

    def show_converter(self) -> None:
        self._clear()
        self.root.geometry("820x650")
        ExcelForTallyApp(
            self.root,
            on_logout=self.show_login,
            on_change_profile=self.show_profiles,
        )


def main() -> None:
    root = tk.Tk()
    LoginFlow(root)
    root.mainloop()
