"""Shared responsive layout, theme, and zoom controls for the desktop UI."""

from __future__ import annotations

import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont

from PIL import Image, ImageTk


THEMES = {
    "light": {
        "bg": "#f3f6f8", "surface": "#ffffff", "fg": "#23313a", "muted": "#52636d",
        "title": "#123d54", "accent": "#145b7a", "accent_fg": "#ffffff", "border": "#d6e0e5",
    },
    "dark": {
        "bg": "#17212b", "surface": "#202d39", "fg": "#e6edf2", "muted": "#adbbc5",
        "title": "#8ed6f3", "accent": "#176b8e", "accent_fg": "#ffffff", "border": "#3a4a57",
    },
}
THEME_OPTIONS = ("System default", "Light", "Dark")
THEME_TO_MODE = {"System default": "system", "Light": "light", "Dark": "dark"}


def _preferences_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return base / "ExcelForTally" / "preferences.json"


def _system_theme() -> str:
    if os.name == "nt":
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            return "light" if value else "dark"
        except (OSError, ImportError):
            pass
    return "light"


def responsive_geometry(root: tk.Tk, width: int, height: int) -> None:
    screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
    actual_w, actual_h = min(width, max(360, screen_w - 48)), min(height, max(320, screen_h - 80))
    root.geometry(f"{actual_w}x{actual_h}")
    root.minsize(min(420, actual_w), min(350, actual_h))
    root.resizable(True, True)


def scrollable_page(root: tk.Tk, background: str) -> ttk.Frame:
    host = ttk.Frame(root)
    host.pack(fill="both", expand=True)
    canvas = tk.Canvas(host, highlightthickness=0, background=background)
    root._scroll_canvas = canvas
    scrollbar = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    content = ttk.Frame(canvas)
    window = canvas.create_window((0, 0), window=content, anchor="nw")

    def update_scroll_region(_event=None) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_content(event) -> None:
        canvas.itemconfigure(window, width=event.width)

        def resize_wrapped(widget: tk.Widget) -> None:
            try:
                wraplength = int(widget.cget("wraplength"))
                if wraplength:
                    widget.configure(wraplength=max(180, event.width - 72))
            except (tk.TclError, TypeError, ValueError):
                pass
            for child in widget.winfo_children():
                resize_wrapped(child)

        resize_wrapped(content)

    content.bind("<Configure>", update_scroll_region)
    canvas.bind("<Configure>", resize_content)
    canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"))
    return content


class Appearance:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.path = _preferences_path()
        try:
            saved = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            saved = {}
        self.mode = saved.get("theme", "system")
        if self.mode not in {"system", "light", "dark"}:
            self.mode = "system"
        try:
            self.zoom = float(saved.get("zoom", 1.0))
        except (TypeError, ValueError):
            self.zoom = 1.0
        self.zoom = min(1.5, max(0.8, self.zoom))
        self._fonts: dict[tuple[int, str], tkfont.Font] = {}
        self.style = ttk.Style(root)
        self.logo: ImageTk.PhotoImage | None = None
        self.apply()
        self.root.after(2000, self._poll_system_theme)

    @property
    def current_theme(self) -> str:
        return _system_theme() if self.mode == "system" else self.mode

    @property
    def colors(self) -> dict[str, str]:
        return THEMES[self.current_theme]

    def font(self, size: int, weight: str = "normal") -> tkfont.Font:
        key = (size, weight)
        if key not in self._fonts:
            self._fonts[key] = tkfont.Font(
                root=self.root,
                family="Segoe UI",
                size=round(size * self.zoom),
                weight=weight,
            )
        return self._fonts[key]

    def _save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            temp = self.path.with_suffix(".tmp")
            temp.write_text(json.dumps({"theme": self.mode, "zoom": self.zoom}, indent=2), encoding="utf-8")
            temp.replace(self.path)
        except OSError:
            pass

    def apply(self) -> None:
        c = self.colors
        self.root.configure(bg=c["bg"])
        canvas = getattr(self.root, "_scroll_canvas", None)
        if canvas is not None and canvas.winfo_exists():
            canvas.configure(background=c["bg"])
        for (size, weight), font in self._fonts.items():
            font.configure(size=round(size * self.zoom))
        self.style.theme_use("clam")
        self.style.configure("TFrame", background=c["bg"])
        self.style.configure("TLabel", background=c["bg"], foreground=c["fg"], font=self.font(10))
        self.style.configure("Title.TLabel", font=self.font(20, "bold"), foreground=c["title"])
        self.style.configure("BrandName.TLabel", font=self.font(16, "bold"), foreground=c["title"])
        self.style.configure("Sub.TLabel", background=c["bg"], foreground=c["muted"], font=self.font(9))
        self.style.configure("Error.TLabel", background=c["bg"], foreground="#ff6464", font=self.font(10, "bold"))
        self.style.configure("Card.TFrame", background=c["surface"], relief="solid", borderwidth=1)
        self.style.configure("Card.TLabel", background=c["surface"], foreground=c["fg"], font=self.font(10))
        self.style.configure("Path.TLabel", background=c["surface"], foreground=c["muted"], font=self.font(9))
        self.style.configure("TEntry", fieldbackground=c["surface"], foreground=c["fg"], insertcolor=c["fg"], padding=(8, 7), font=self.font(11))
        self.style.configure("TCombobox", fieldbackground=c["surface"], background=c["surface"], foreground=c["fg"])
        self.style.map("TCombobox", fieldbackground=[("readonly", c["surface"])], foreground=[("readonly", c["fg"])])
        self.style.configure("TCheckbutton", background=c["bg"], foreground=c["fg"], font=self.font(10))
        self.style.map("TCheckbutton", background=[("active", c["bg"])], foreground=[("active", c["fg"])])
        self.style.configure("TButton", font=self.font(10), padding=(9, 6), background=c["surface"], foreground=c["fg"])
        self.style.map("TButton", background=[("active", c["border"])])
        self.style.configure("Accent.TButton", font=self.font(10, "bold"), padding=(12, 8), background=c["accent"], foreground=c["accent_fg"])
        self.style.map("Accent.TButton", background=[("active", c["title"])], foreground=[("active", c["accent_fg"])])

    def set_mode(self, option: str) -> None:
        self.mode = THEME_TO_MODE.get(option, "system")
        self._save()
        self.apply()

    def zoom_by(self, amount: float) -> None:
        self.zoom = min(1.5, max(0.8, round(self.zoom + amount, 2)))
        self._save()
        self.apply()

    def option_name(self) -> str:
        return {"system": "System default", "light": "Light", "dark": "Dark"}[self.mode]

    def controls(self, parent: ttk.Frame) -> ttk.Frame:
        controls = ttk.Frame(parent)
        ttk.Button(controls, text="A−", width=3, command=lambda: self.zoom_by(-0.1)).pack(side="left", padx=2)
        ttk.Button(controls, text="A+", width=3, command=lambda: self.zoom_by(0.1)).pack(side="left", padx=2)
        theme = ttk.Combobox(controls, state="readonly", width=15, values=THEME_OPTIONS)
        theme.set(self.option_name())
        theme.bind("<<ComboboxSelected>>", lambda _event: self.set_mode(theme.get()))
        theme.pack(side="left", padx=(6, 0))
        return controls

    def brand_header(self, parent: ttk.Frame) -> ttk.Frame:
        header = ttk.Frame(parent)
        header.pack(fill="x", pady=(0, 10))
        brand = ttk.Frame(header)
        brand.pack(fill="x", expand=True)
        try:
            logo_path = Path(__file__).parent / "data" / "revtech_logo.png"
            image = Image.open(logo_path).convert("RGBA")
            image.thumbnail((58, 58), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(image)
            ttk.Label(brand, image=self.logo).pack(side="left", padx=(0, 10))
        except (OSError, tk.TclError):
            self.logo = None
        titles = ttk.Frame(brand)
        titles.pack(side="left", fill="x", expand=True)
        ttk.Label(titles, text="Tally Input Generator", style="BrandName.TLabel").pack(anchor="w")
        ttk.Label(titles, text="Rev Tech Solutions", style="Sub.TLabel").pack(anchor="w")
        self.controls(header).pack(anchor="e", pady=(8, 0))
        return header

    def _poll_system_theme(self) -> None:
        if self.mode == "system":
            self.apply()
        if self.root.winfo_exists():
            self.root.after(2000, self._poll_system_theme)
