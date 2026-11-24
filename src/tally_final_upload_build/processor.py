"""Convert hospital reports into the sample Tally import workbook layout."""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
from openpyxl.styles import Font, PatternFill

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


SALES_HEADERS = [
    "InvNo", "Inv_Dt", "Pty_Name", "Vch_Type", "GSTIN", "StateOfSupply",
    "Product_Name", "HSNCode", "Qty", "UOM", "TaxPer", "TaxableAmt",
    "IGSTAmt", "SGSTAmt", "CGSTAmt", "Cess", "Other Amt", "Net_Amt",
    "Narration ", "Discount", "Sales Ledger",
]
VOUCHER_HEADERS = ["Date", "Rec No", "Bank / Cash", "Credit Account", "Amount", "Cheque No", "Chk Date", "Narration"]
PAYMENT_HEADERS = ["Date", "Payment No", "Bank / Cash", "Debit Account", "Amount", "Cheque No", "Chk Date", "Narration"]
JOURNAL_HEADERS = ["VCH Type", "Bill Number", "Bill Date", "Doctor", "Debit Amt", "Item Value", "Narration"]
DISCOUNT_HEADERS = JOURNAL_HEADERS + [""] * 9
DOCTOR_HEADERS = JOURNAL_HEADERS + [""]
PAYTM_HEADERS = VOUCHER_HEADERS + ["", ""]


def _clean_header(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def _rows_from_upload(filename: str, payload: bytes) -> tuple[list[str], list[list[Any]]]:
    """Read the first worksheet from legacy .xls or modern .xlsx files."""
    suffix = Path(filename).suffix.casefold()
    if suffix == ".xls":
        book = xlrd.open_workbook(file_contents=payload, on_demand=True)
        sheet = book.sheet_by_index(0)
        rows: list[list[Any]] = []
        for row_index in range(sheet.nrows):
            values: list[Any] = []
            for col_index in range(sheet.ncols):
                cell = sheet.cell(row_index, col_index)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate.xldate_as_datetime(value, book.datemode)
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    value = None
                values.append(value)
            rows.append(values)
        book.release_resources()
    elif suffix == ".xlsx":
        book = load_workbook(BytesIO(payload), read_only=True, data_only=True)
        sheet = book.active
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        book.close()
    else:
        raise ValueError(f"{filename}: expected an .xls or .xlsx file")

    header_index = next(
        (i for i, row in enumerate(rows[:30]) if sum(v is not None and str(v).strip() != "" for v in row) >= 3),
        None,
    )
    if header_index is None:
        raise ValueError(f"{filename}: could not find a header row")
    headers = [str(value or "").strip() for value in rows[header_index]]
    return headers, rows[header_index + 1 :]


def _column(headers: list[str], name: str, *aliases: str) -> int:
    normalized = {_clean_header(header): i for i, header in enumerate(headers) if header}
    for candidate in (name, *aliases):
        key = _clean_header(candidate)
        if key in normalized:
            return normalized[key]
    raise ValueError(f"Required column not found: {name}")


def _get(row: list[Any], headers: list[str], name: str, *aliases: str, default: Any = None) -> Any:
    try:
        index = _column(headers, name, *aliases)
    except ValueError:
        return default
    return row[index] if index < len(row) else default


def _has_values(row: list[Any]) -> bool:
    return any(value is not None and str(value).strip() for value in row)


def _party(visit_no: Any, name: Any) -> str:
    visit = str(visit_no or "").strip()
    patient = str(name or "").strip()
    if visit and patient:
        return f"{visit}-{patient}"
    return visit or patient


def _date(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                pass
    return value


def _datetime(value: Any) -> Any:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
    return value


def _amount(value: Any) -> Any:
    if value in (None, ""):
        return 0
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return value


def _bank(mode: Any) -> str:
    text = str(mode or "").strip().upper()
    if text in {"UPI", "CARD"}:
        return "UBI"
    if text == "EFT":
        return "Indian Overseas Bank C.A 3762"
    return text


def sales_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, rows = _rows_from_upload(filename, payload)
    bill = _column(headers, "Bill Number")
    bill_date = _column(headers, "Bill Date")
    visit = _column(headers, "Visit Number")
    patient = _column(headers, "Patient Name")
    service = _column(headers, "Services / Charge Name")
    item_value = _column(headers, "Item Value")
    discount = _column(headers, "Discount")
    result = []
    for row in rows:
        if not _has_values(row):
            continue
        party = _party(row[visit], row[patient])
        amount = _amount(row[item_value])
        result.append([
            row[bill], _date(row[bill_date]), party, "Auto Sales", "", "Tamil Nadu",
            "", "", "", "", 0, amount, 0, 0, 0, 0, 0, amount,
            party, _amount(row[discount]), row[service],
        ])
    return result


def receipt_rows(filename: str, payload: bytes, *, advance: bool = False) -> list[list[Any]]:
    headers, rows = _rows_from_upload(filename, payload)
    number = _column(headers, "Receipt Number")
    date_col = _column(headers, "Receipt Date")
    visit = _column(headers, "Visit Number")
    patient = _column(headers, "Patient Name")
    mode = _column(headers, "Mode")
    cheque = _column(headers, "DD / Cheque Info")
    amount = _column(headers, "Receipt Value")
    result = []
    for row in rows:
        if not _has_values(row):
            continue
        party = _party(row[visit], row[patient])
        result.append([
            _date(row[date_col]), row[number], _bank(row[mode]), party,
            _amount(row[amount]), row[cheque], "", party,
        ])
    return result


def refund_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, rows = _rows_from_upload(filename, payload)
    result = []
    for row in rows:
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit No"), _get(row, headers, "Patient Name"))
        result.append([
            _date(_get(row, headers, "Payment Date")), _get(row, headers, "Payment Number"),
            _bank(_get(row, headers, "Mode")), party,
            _amount(_get(row, headers, "Payment Value")),
            _get(row, headers, "Cheque No."), "", party,
        ])
    return result


def discount_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, rows = _rows_from_upload(filename, payload)
    result = []
    for row in rows:
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        bill = _get(row, headers, "Bill No")
        bill_date = _datetime(_get(row, headers, "Bill Time"))
        value = _amount(_get(row, headers, "Discount Value"))
        result.append(["Auto Journal", bill, bill_date, "Discount Allowed", value, "", party])
        result.append(["Auto Journal", bill, bill_date, party, "", value, party])
    return result


def doctor_fee_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, rows = _rows_from_upload(filename, payload)
    result = []
    for row in rows:
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        bill_date = _datetime(_get(row, headers, "Bill Date"))
        bill = _get(row, headers, "Bill Number")
        value = _amount(_get(row, headers, "Net Value"))
        doctor = _get(row, headers, "Doctor")
        result.append(["Auto Journal", bill, bill_date, "Professional Charges", value, "", party])
        result.append(["Auto Journal", bill, bill_date, doctor, "", value, party])
    return result


def make_workbook(tables: dict[str, tuple[list[str], Iterable[list[Any]]]]) -> bytes:
    book = Workbook()
    book.remove(book.active)
    for title, (headers, rows) in tables.items():
        sheet = book.create_sheet(title)
        sheet.append(headers)
        for row in rows:
            sheet.append(row + [""] * max(0, len(headers) - len(row)))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="245A73")
        for column in sheet.columns:
            values = [len(str(cell.value)) for cell in list(column)[:300] if cell.value is not None]
            sheet.column_dimensions[get_column_letter(column[0].column)].width = min(max(max(values, default=10) + 2, 12), 40)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (date, datetime)):
                    if title in {"Advances", "Receipt"}:
                        cell.number_format = "dd/mm/yy"
                    elif title in {"Discount Allowed", "Doctor Fees"}:
                        cell.number_format = "yyyy-mm-dd hh:mm:ss.000"
                    elif title == "Paytm":
                        cell.number_format = "yyyy-mm-dd"
                    else:
                        cell.number_format = "dd/mm/yyyy"
    output = BytesIO()
    book.save(output)
    return output.getvalue()
