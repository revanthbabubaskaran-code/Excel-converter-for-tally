"""Small license-gated installer bootstrapper for the standalone application."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tally_final_upload_build.licensing import (
    DEFAULT_LICENSE_SERVICE_URL,
    LicenseError,
    activate_license,
    check_license_service,
    check_online_license_request,
    is_licensed,
    request_code,
    submit_online_license_request,
    verify_license,
)


APP_NAME = "Excel for Tally"
APP_EXE_NAME = "Excel for Tally.exe"
INSTALL_DIRECTORY = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Programs" / APP_NAME


class SetupWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} Setup")
        self.root.geometry("660x700")
        self.root.minsize(600, 650)
        self.root.configure(bg="#f3f6f8")
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        self.install_directory = INSTALL_DIRECTORY
        self._service_after_id: str | None = None
        self._retry_delay_ms = 5_000
        self._install_check_running = False

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f6f8")
        style.configure("TLabel", background="#f3f6f8", foreground="#23313a", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground="#123d54")
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#52636d")
        style.configure("Error.TLabel", foreground="#c62828", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8))
        style.configure("TEntry", padding=(8, 7), font=("Segoe UI", 10))

        outer = ttk.Frame(root, padding=(30, 24, 30, 22))
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text=f"Install {APP_NAME}", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Setup requests approval for this Windows device automatically. You can copy its device token and send it to the software owner.",
            style="Sub.TLabel",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(7, 16))

        self.error_label = ttk.Label(outer, text="", style="Error.TLabel", wraplength=520)
        self.error_label.pack(anchor="w", pady=(0, 5))
        try:
            self.device_request_code = request_code()
        except LicenseError as exc:
            self.device_request_code = ""
            self.error_label.configure(text=str(exc))

        ttk.Label(outer, text="Device request token (send this to the software owner)").pack(anchor="w")
        token_row = ttk.Frame(outer)
        token_row.pack(fill="x", pady=(4, 12))
        self.request_token = tk.StringVar(value=self.device_request_code)
        self.request_entry = ttk.Entry(token_row, textvariable=self.request_token, state="readonly", font=("Consolas", 10))
        self.request_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.copy_token_button = ttk.Button(token_row, text="Copy token", command=self.copy_request_token)
        self.copy_token_button.pack(side="right")

        ttk.Label(outer, text="Owner-issued device license").pack(anchor="w")
        self.license_entry = ttk.Entry(outer, font=("Consolas", 9))
        self.license_entry.pack(fill="x", pady=(4, 7))

        self.status = ttk.Label(
            outer,
            text="Enter the owner-issued license for this device. Setup checks it before installing.",
            style="Sub.TLabel",
            wraplength=520,
        )
        self.status.pack(anchor="w", pady=(2, 8))

        ttk.Label(outer, text="Installation folder").pack(anchor="w", pady=(2, 2))
        install_path_row = ttk.Frame(outer)
        install_path_row.pack(fill="x", pady=(2, 12))
        self.install_path = tk.StringVar(value=str(self.install_directory))
        self.install_path_entry = ttk.Entry(install_path_row, textvariable=self.install_path, state="readonly")
        self.install_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.browse_button = ttk.Button(install_path_row, text="Browse…", command=self.choose_install_directory)
        self.browse_button.pack(side="right")

        self.install_button = ttk.Button(
            outer,
            text="Check license and install",
            style="Accent.TButton",
            command=self.install,
        )
        self.install_button.pack(anchor="e", pady=(2, 14))

        try:
            if is_licensed():
                self.status.configure(
                    text="A valid license is already saved for this device. Click Install to continue; no license entry is needed."
                )
                self.license_entry.configure(state="disabled")
                self.install_button.configure(text="Install using saved license")
            elif self.device_request_code:
                self.service_call()
        except LicenseError as exc:
            self.error_label.configure(text=str(exc))

    def copy_request_token(self) -> None:
        if not self.device_request_code:
            self.status.configure(text="The device token is unavailable on this computer.", style="Error.TLabel")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.device_request_code)
        self.root.update()
        self.status.configure(text="Device token copied. Send it to the software owner.", style="Sub.TLabel")

    def choose_install_directory(self) -> None:
        initial_directory = self.install_directory
        while not initial_directory.is_dir() and initial_directory != initial_directory.parent:
            initial_directory = initial_directory.parent
        selected = filedialog.askdirectory(
            parent=self.root,
            title="Select installation folder",
            initialdir=str(initial_directory),
            mustexist=True,
        )
        if selected:
            self.install_directory = Path(selected).expanduser()
            self.install_path.set(str(self.install_directory))

    def service_call(self) -> None:
        self._service_after_id = None
        code = self.device_request_code
        if not code:
            self.status.configure(text="Could not create a license request for this device.", style="Error.TLabel")
            return
        self.install_button.configure(state="disabled")
        self.status.configure(text="Sending this device's license request…", style="Sub.TLabel")

        def worker() -> None:
            result: dict | None = None
            error = ""
            try:
                result = submit_online_license_request(code, DEFAULT_LICENSE_SERVICE_URL)
            except LicenseError as exc:
                error = str(exc)
            try:
                self.root.after(0, lambda: self.finish_service_call(result, error))
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def finish_service_call(self, result: dict | None, error: str) -> None:
        self.install_button.configure(state="normal")
        if error:
            self.status.configure(
                text=f"Could not send the license request: {error} Retrying automatically in {self._retry_delay_ms // 1_000} seconds. Keep setup open.",
                style="Error.TLabel",
            )
            self._schedule_service_check(self._next_retry_delay(), submit=True)
            return
        if result is None:
            self.status.configure(
                text=f"The license service returned no result. Retrying automatically in {self._retry_delay_ms // 1_000} seconds. Keep setup open.",
                style="Error.TLabel",
            )
            self._schedule_service_check(self._next_retry_delay(), submit=True)
            return
        self._retry_delay_ms = 5_000
        state = result.get("status")
        if state == "approved" and result.get("license_token"):
            self._receive_license(str(result["license_token"]))
        elif state == "pending":
            self.status.configure(
                text="License request sent. Waiting for owner approval; setup will check automatically.",
                style="Sub.TLabel",
            )
            self._schedule_service_check(5_000)
        elif state == "rejected":
            self.status.configure(text="The owner declined this license request.", style="Error.TLabel")
        else:
            self.status.configure(
                text=f"The license service returned an unexpected status. Retrying automatically in {self._retry_delay_ms // 1_000} seconds. Keep setup open.",
                style="Error.TLabel",
            )
            self._schedule_service_check(self._next_retry_delay(), submit=True)

    def _next_retry_delay(self) -> int:
        delay = self._retry_delay_ms
        self._retry_delay_ms = min(self._retry_delay_ms * 2, 60_000)
        return delay

    def _schedule_service_check(self, delay_ms: int, *, submit: bool = False) -> None:
        if self._service_after_id is not None:
            try:
                self.root.after_cancel(self._service_after_id)
            except tk.TclError:
                pass
        callback = self.service_call if submit else self.check_service_status
        self._service_after_id = self.root.after(delay_ms, callback)

    def check_service_status(self) -> None:
        self._service_after_id = None
        code = self.device_request_code
        if not code:
            return
        self.status.configure(text="Checking for owner approval…", style="Sub.TLabel")

        def worker() -> None:
            result: dict | None = None
            error = ""
            try:
                result = check_online_license_request(code, DEFAULT_LICENSE_SERVICE_URL)
            except LicenseError as exc:
                error = str(exc)
            try:
                self.root.after(0, lambda: self.finish_status_check(result, error))
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def finish_status_check(self, result: dict | None, error: str) -> None:
        if error or result is None:
            detail = error or "The license service returned no result."
            delay = self._next_retry_delay()
            self.status.configure(
                text=f"Could not check approval: {detail} Retrying automatically in {delay // 1_000} seconds. Keep setup open.",
                style="Error.TLabel",
            )
            self._schedule_service_check(delay)
            return
        self._retry_delay_ms = 5_000
        state = result.get("status")
        if state == "approved" and result.get("license_token"):
            self._receive_license(str(result["license_token"]))
        elif state == "pending":
            self.status.configure(
                text="License request sent. Waiting for owner approval; setup will check automatically.",
                style="Sub.TLabel",
            )
            self._schedule_service_check(5_000)
        elif state == "rejected":
            self.status.configure(text="The owner declined this license request.", style="Error.TLabel")
        else:
            delay = self._next_retry_delay()
            self.status.configure(
                text=f"The license service returned an unexpected status. Retrying automatically in {delay // 1_000} seconds. Keep setup open.",
                style="Error.TLabel",
            )
            self._schedule_service_check(delay)

    def _receive_license(self, token: str) -> None:
        if self._service_after_id is not None:
            try:
                self.root.after_cancel(self._service_after_id)
            except tk.TclError:
                pass
            self._service_after_id = None
        self.license_entry.configure(state="normal")
        self.license_entry.delete(0, tk.END)
        self.license_entry.insert(0, token)
        self.status.configure(
            text="Owner approved this device. License received; click Check license and install to continue.",
            style="Sub.TLabel",
        )

    @staticmethod
    def payload_path() -> Path:
        bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        return bundle_root / "payload" / APP_EXE_NAME

    def install(self) -> None:
        if self._install_check_running:
            return
        self._install_check_running = True
        self.install_button.configure(state="disabled")
        self.status.configure(text="Checking the internet connection to the licensing service…", style="Sub.TLabel")

        def worker() -> None:
            error = ""
            try:
                check_license_service(DEFAULT_LICENSE_SERVICE_URL)
            except LicenseError as exc:
                error = str(exc)
            try:
                self.root.after(0, lambda: self.finish_network_check(error))
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def finish_network_check(self, error: str) -> None:
        self._install_check_running = False
        self.install_button.configure(state="normal")
        if error:
            self.status.configure(
                text="Setup cannot reach the online licensing service. Connect this computer to the internet, then try again.",
                style="Error.TLabel",
            )
            messagebox.showwarning(
                "Internet connection required",
                "Connect this computer to the internet. Setup must reach the licensing service before it can install.",
                parent=self.root,
            )
            return
        self._install_after_network_check()

    def _install_after_network_check(self) -> None:
        try:
            already_licensed = is_licensed()
            token = None if already_licensed else self.license_entry.get().strip()
            if not already_licensed and not token:
                self.status.configure(
                    text="License request sent. Waiting for owner approval; setup will check automatically. Keep setup open.",
                    style="Sub.TLabel",
                )
                return
            if not already_licensed and not verify_license(token or ""):
                self.status.configure(
                    text="License is incorrect or belongs to a different device. Nothing was installed.",
                    style="Error.TLabel",
                )
                messagebox.showerror(
                    "License verification failed",
                    "This owner-issued license is invalid or was issued for another device. The software was not installed.",
                    parent=self.root,
                )
                return

            payload = self.payload_path()
            if not payload.is_file():
                raise OSError("The application files are missing from this setup package.")
            if self.install_directory.exists() and not self.install_directory.is_dir():
                raise OSError("The selected installation path is not a folder.")

            self.install_button.configure(state="disabled")
            self.status.configure(
                text=(
                    "Installing with the license already saved on this device…"
                    if already_licensed
                    else "License verified. Installing…"
                ),
                style="Sub.TLabel",
            )
            self.install_directory.mkdir(parents=True, exist_ok=True)
            destination = self.install_directory / APP_EXE_NAME
            shutil.copy2(payload, destination)
            if token:
                activate_license(token)

            completion = (
                "Installation completed using the license already saved on this device."
                if already_licensed
                else "Installation completed. Your device license is saved."
            )
            self.status.configure(text=completion, style="Sub.TLabel")
            messagebox.showinfo(
                "Installation complete",
                f"{APP_NAME} was installed. {completion}\n\nRun:\n{destination}",
                parent=self.root,
            )
            self.root.destroy()
        except (OSError, LicenseError) as exc:
            self.install_button.configure(state="normal")
            self.status.configure(text=f"Installation failed: {exc}", style="Error.TLabel")
            messagebox.showerror("Installation failed", str(exc), parent=self.root)


def main() -> None:
    root = tk.Tk()
    SetupWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
