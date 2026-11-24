"""Native Windows desktop interface for the report converter."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from collections.abc import Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tally_final_upload_build.processor import (
    DISCOUNT_HEADERS,
    DOCTOR_HEADERS,
    PAYMENT_HEADERS,
    PAYTM_HEADERS,
    SALES_HEADERS,
    VOUCHER_HEADERS,
    discount_rows,
    doctor_fee_rows,
    make_workbook,
    receipt_rows,
    refund_rows,
    sales_rows,
)


REPORTS = {
    "charge": ("Charge Item Report", "charge item"),
    "doctor": ("Doctor Fees Report", "doctor fee"),
    "advance": ("IP Advance Report", "ip advance"),
    "discount": ("IP Discount Report", "ip discount"),
    "receipt": ("IP Receipt Report", "ip receipt"),
    "refund": ("IP Refund Report", "ip refund"),
}


class ExcelForTallyApp:
    def __init__(self, root: tk.Tk, on_logout: Callable[[], None] | None = None) -> None:
        self.root = root
        self.on_logout = on_logout
        self.root.title("Excel for Tally")
        self.root.geometry("820x650")
        self.root.minsize(720, 560)
        self.root.configure(bg="#f3f6f8")
        self.paths: dict[str, Path] = {}
        self.file_labels: dict[str, ttk.Label] = {}

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f6f8")
        style.configure("TLabel", background="#f3f6f8", foreground="#23313a", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#123d54")
        style.configure("Sub.TLabel", font=("Segoe UI", 10), foreground="#52636d")
        style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("Card.TLabel", background="#ffffff", foreground="#23313a", font=("Segoe UI", 10))
        style.configure("Path.TLabel", background="#ffffff", foreground="#53666f", font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9))

        outer = ttk.Frame(root, padding=(28, 24, 28, 22))
        outer.pack(fill="both", expand=True)

        heading = ttk.Frame(outer)
        heading.pack(fill="x")
        ttk.Label(heading, text="Excel for Tally", style="Title.TLabel").pack(side="left", anchor="w")
        if self.on_logout is not None:
            ttk.Button(heading, text="Log out", command=self.on_logout).pack(side="right", anchor="e")
        ttk.Label(
            outer,
            text="Choose your hospital reports and create the formatted Excel workbook on this computer.",
            style="Sub.TLabel",
            wraplength=760,
        ).pack(anchor="w", pady=(5, 18))

        note = ttk.Frame(outer, style="Card.TFrame", padding=13)
        note.pack(fill="x", pady=(0, 16))
        ttk.Label(
            note,
            text=(
                "The PDF bank statement is ignored. The sample's Paytm tab has 662 OP receipt rows, "
                "but those records are not present in the supplied Excel reports; the generated Paytm "
                "tab will contain its headers only."
            ),
            style="Card.TLabel",
            wraplength=720,
            justify="left",
        ).pack(anchor="w")

        ttk.Label(outer, text="Required Excel reports", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        reports_card = ttk.Frame(outer, style="Card.TFrame", padding=(14, 8, 14, 10))
        reports_card.pack(fill="x")

        for key, (title, _) in REPORTS.items():
            row = ttk.Frame(reports_card, style="Card.TFrame")
            row.pack(fill="x", pady=5)
            ttk.Label(row, text=title, style="Card.TLabel", width=22).pack(side="left", anchor="w")
            file_label = ttk.Label(row, text="No file selected", style="Path.TLabel", width=43, anchor="w")
            file_label.pack(side="left", fill="x", expand=True, padx=(8, 10))
            self.file_labels[key] = file_label
            ttk.Button(row, text="Browse…", command=lambda report_key=key: self.choose_file(report_key)).pack(side="right")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(15, 0))
        ttk.Label(
            footer,
            text="IP Bill Report is not used because the sample output has no matching sheet.",
            style="Sub.TLabel",
        ).pack(anchor="w")
        ttk.Label(footer, text="All file processing happens locally on this computer.", style="Sub.TLabel").pack(anchor="w", pady=(3, 0))

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", side="bottom", pady=(18, 0))
        self.status = ttk.Label(bottom, text="Select all six reports to continue.", style="Sub.TLabel")
        self.status.pack(side="left", fill="x", expand=True)
        self.generate_button = ttk.Button(bottom, text="Create output Excel", style="Accent.TButton", command=self.create_output)
        self.generate_button.pack(side="right")
        self.update_ready_state()

    def choose_file(self, key: str) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title=f"Choose {REPORTS[key][0]}",
            filetypes=[("Excel workbooks", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.paths[key] = Path(path)
        self.file_labels[key].configure(text=self.paths[key].name)
        self.update_ready_state()

    def update_ready_state(self) -> None:
        missing = [REPORTS[key][0] for key in REPORTS if key not in self.paths]
        self.generate_button.configure(state="normal" if not missing else "disabled")
        if missing:
            self.status.configure(text=f"Still needed: {', '.join(missing)}")
        else:
            self.status.configure(text="All required reports are ready.")

    def create_output(self) -> None:
        suggested = f"Tally_Final_Output_{date.today():%Y%m%d}.xlsx"
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save output workbook",
            initialfile=suggested,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output_path:
            return

        self.generate_button.configure(state="disabled")
        self.status.configure(text="Creating workbook…")
        self.root.update_idletasks()
        try:
            def read(key: str) -> tuple[str, bytes]:
                path = self.paths[key]
                return path.name, path.read_bytes()

            charge_name, charge_bytes = read("charge")
            doctor_name, doctor_bytes = read("doctor")
            advance_name, advance_bytes = read("advance")
            discount_name, discount_bytes = read("discount")
            receipt_name, receipt_bytes = read("receipt")
            refund_name, refund_bytes = read("refund")

            tables = {
                "Auto Sales": (SALES_HEADERS, sales_rows(charge_name, charge_bytes)),
                "Advances": (VOUCHER_HEADERS, receipt_rows(advance_name, advance_bytes)),
                "Receipt": (VOUCHER_HEADERS, receipt_rows(receipt_name, receipt_bytes)),
                "Discount Allowed": (DISCOUNT_HEADERS, discount_rows(discount_name, discount_bytes)),
                "Doctor Fees": (DOCTOR_HEADERS, doctor_fee_rows(doctor_name, doctor_bytes)),
                "Auto Payment": (PAYMENT_HEADERS, refund_rows(refund_name, refund_bytes)),
                "Paytm": (PAYTM_HEADERS, []),
                "Formula": ([""] * 13, []),
            }
            output = make_workbook(tables)
            Path(output_path).write_bytes(output)
            summary = "\n".join(f"{name}: {len(rows)} rows" for name, (_, rows) in tables.items())
            self.status.configure(text=f"Saved: {Path(output_path).name}")
            messagebox.showinfo(
                "Workbook created",
                f"The output workbook was saved to:\n{output_path}\n\n{summary}",
                parent=self.root,
            )
        except Exception as exc:
            self.status.configure(text="Could not create the workbook.")
            messagebox.showerror("Workbook creation failed", str(exc), parent=self.root)
        finally:
            self.update_ready_state()


def main() -> None:
    from tally_final_upload_build.login import main as login_main

    login_main()


if __name__ == "__main__":
    main()
