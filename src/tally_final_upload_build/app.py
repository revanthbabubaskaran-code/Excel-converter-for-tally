"""Native Windows desktop interface for the report converter."""

from __future__ import annotations

from pathlib import Path
from collections.abc import Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tally_final_upload_build.processor import (
    DISCOUNT_HEADERS,
    DOCTOR_HEADERS,
    FORMULA_HEADERS,
    PAYMENT_HEADERS,
    PAYTM_HEADERS,
    SALES_HEADERS,
    VOUCHER_HEADERS,
    discount_rows,
    doctor_fee_rows,
    formula_rows,
    make_workbook,
    paytm_rows,
    receipt_rows,
    refund_rows,
    sales_rows,
)


REPORTS = {
    "charge": ("Charge Item Report", "charge item"),
    "bill": ("IP Bill Report", "ip bill"),
    "doctor": ("Doctor Fees Report", "doctor fee"),
    "advance": ("IP Advance Report", "ip advance"),
    "discount": ("IP Discount Report", "ip discount"),
    "receipt": ("IP Receipt Report", "ip receipt"),
    "refund": ("IP Refund Report", "ip refund"),
    "paytm": ("Card & UPI Payments", "card upi payment"),
}


class ExcelForTallyApp:
    def __init__(self, root: tk.Tk, on_logout: Callable[[], None] | None = None) -> None:
        self.root = root
        self.on_logout = on_logout
        self.root.title("Excel for Tally")
        self.root.geometry("850x735")
        self.root.minsize(760, 680)
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

        outer = ttk.Frame(root, padding=(28, 22, 28, 20))
        outer.pack(fill="both", expand=True)
        heading = ttk.Frame(outer)
        heading.pack(fill="x")
        ttk.Label(heading, text="Excel for Tally", style="Title.TLabel").pack(side="left", anchor="w")
        if self.on_logout is not None:
            ttk.Button(heading, text="Log out", command=self.on_logout).pack(side="right", anchor="e")
        ttk.Label(
            outer,
            text="Choose the hospital Excel reports and create one workbook in the sample format.",
            style="Sub.TLabel",
            wraplength=780,
        ).pack(anchor="w", pady=(5, 16))

        ttk.Label(outer, text="Excel reports (select one or more)", font=("Segoe UI", 12, "bold")).pack(
            anchor="w", pady=(0, 8)
        )
        reports_card = ttk.Frame(outer, style="Card.TFrame", padding=(14, 6, 14, 8))
        reports_card.pack(fill="x")
        for key, (title, _) in REPORTS.items():
            row = ttk.Frame(reports_card, style="Card.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=title, style="Card.TLabel", width=22).pack(side="left", anchor="w")
            file_label = ttk.Label(row, text="No file selected", style="Path.TLabel", width=48, anchor="w")
            file_label.pack(side="left", fill="x", expand=True, padx=(8, 10))
            self.file_labels[key] = file_label
            ttk.Button(row, text="Browse…", command=lambda report_key=key: self.choose_file(report_key)).pack(side="right")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(12, 0))
        ttk.Label(footer, text="All file processing happens locally on this computer.", style="Sub.TLabel").pack(anchor="w")

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", side="bottom", pady=(14, 0))
        self.status = ttk.Label(bottom, text="Select one or more reports to continue.", style="Sub.TLabel")
        self.status.pack(side="left", fill="x", expand=True)
        self.generate_button = ttk.Button(
            bottom, text="Create output workbook", style="Accent.TButton", command=self.create_output
        )
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
        self.generate_button.configure(state="normal" if self.paths else "disabled")
        self.status.configure(
            text=f"{len(self.paths)} of {len(REPORTS)} reports selected."
            if self.paths else "Select one or more reports to continue."
        )

    def create_output(self) -> None:
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save formatted output workbook",
            defaultextension=".xlsx",
            initialfile="Tally_Output.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output_path:
            return
        destination = Path(output_path)
        if destination.exists() and not messagebox.askyesno(
            "Replace existing workbook?",
            f"{destination.name} already exists. Replace it?",
            parent=self.root,
        ):
            self.status.configure(text="Output was not created.")
            return

        self.generate_button.configure(state="disabled")
        self.status.configure(text="Creating workbook…")
        self.root.update_idletasks()
        try:
            def read(key: str) -> tuple[str, bytes]:
                path = self.paths[key]
                return path.name, path.read_bytes()

            tables: dict[str, tuple[list[str], list[list[object]]]] = {
                "Auto Sales": (SALES_HEADERS, []),
                "Advances": (VOUCHER_HEADERS, []),
                "Receipt": (VOUCHER_HEADERS, []),
                "Discount Allowed": (DISCOUNT_HEADERS, []),
                "Doctor Fees": (DOCTOR_HEADERS, []),
                "Auto Payment": (PAYMENT_HEADERS, []),
                "Paytm": (PAYTM_HEADERS, []),
                "Formula": (FORMULA_HEADERS, []),
            }
            row_counts: dict[str, int] = {}
            if "charge" in self.paths:
                name, payload = read("charge")
                rows = sales_rows(name, payload)
                tables["Auto Sales"] = (SALES_HEADERS, rows)
                row_counts["Auto Sales"] = len(rows)
            if "advance" in self.paths:
                name, payload = read("advance")
                rows = receipt_rows(name, payload)
                tables["Advances"] = (VOUCHER_HEADERS, rows)
                row_counts["Advances"] = len(rows)
            if "receipt" in self.paths:
                name, payload = read("receipt")
                rows = receipt_rows(name, payload)
                tables["Receipt"] = (VOUCHER_HEADERS, rows)
                row_counts["Receipt"] = len(rows)
            if "discount" in self.paths:
                name, payload = read("discount")
                rows = discount_rows(name, payload)
                tables["Discount Allowed"] = (DISCOUNT_HEADERS, rows)
                row_counts["Discount Allowed"] = len(rows)
            if "doctor" in self.paths:
                name, payload = read("doctor")
                rows = doctor_fee_rows(name, payload)
                tables["Doctor Fees"] = (DOCTOR_HEADERS, rows)
                row_counts["Doctor Fees"] = len(rows)
            if "refund" in self.paths:
                name, payload = read("refund")
                rows = refund_rows(name, payload)
                tables["Auto Payment"] = (PAYMENT_HEADERS, rows)
                row_counts["Auto Payment"] = len(rows)
            if "paytm" in self.paths:
                name, payload = read("paytm")
                rows = paytm_rows(name, payload)
                tables["Paytm"] = (PAYTM_HEADERS, rows)
                row_counts["Paytm"] = len(rows)
            if "bill" in self.paths:
                name, payload = read("bill")
                rows = formula_rows(name, payload)
                tables["Formula"] = (FORMULA_HEADERS, rows)
                row_counts["Formula"] = len(rows)

            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(make_workbook(tables))
            summary = "\n".join(f"{sheet}: {count} rows" for sheet, count in row_counts.items())
            if not summary:
                summary = "No source rows were found in the selected reports."
            self.status.configure(text=f"Saved workbook: {destination}")
            messagebox.showinfo(
                "Output workbook created",
                f"The formatted workbook was saved to:\n{destination}\n\n{summary}",
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
