"""Convert hospital reports into the supplied Tally workbook layout."""

from __future__ import annotations

from collections.abc import Iterable
from copy import copy
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import re
from typing import Any

import xlrd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries


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
DOCTOR_HEADERS = JOURNAL_HEADERS + [""] * 2
PAYTM_HEADERS = VOUCHER_HEADERS + ["", ""]
FORMULA_HEADERS = [""] * 13

SHEET_HEADERS = {
    "Auto Sales": SALES_HEADERS,
    "Advances": VOUCHER_HEADERS,
    "Receipt": VOUCHER_HEADERS,
    "Discount Allowed": DISCOUNT_HEADERS,
    "Doctor Fees": DOCTOR_HEADERS,
    "Auto Payment": PAYMENT_HEADERS,
    "Paytm": PAYTM_HEADERS,
    "Formula": FORMULA_HEADERS,
}
LAYOUT_TEMPLATE = Path(__file__).resolve().parent / "data" / "output_layout_template.xlsx"


def _clean_header(value: Any) -> str:
    text = str(value or "").strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


def _rows_from_upload(filename: str, payload: bytes) -> tuple[list[str], list[list[Any]]]:
    """Read the first sheet of a legacy .xls or modern .xlsx workbook."""
    suffix = Path(filename).suffix.casefold()
    if suffix == ".xls":
        book = xlrd.open_workbook(file_contents=payload, on_demand=True)
        try:
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
        finally:
            book.release_resources()
    elif suffix == ".xlsx":
        book = load_workbook(BytesIO(payload), read_only=True, data_only=True)
        try:
            sheet = book.active
            rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        finally:
            book.close()
    else:
        raise ValueError(f"{filename}: expected an .xls or .xlsx file")

    header_index = next(
        (index for index, row in enumerate(rows[:30]) if sum(bool(v is not None and str(v).strip()) for v in row) >= 3),
        None,
    )
    if header_index is None:
        raise ValueError(f"{filename}: could not find a header row")
    headers = [str(value or "").strip() for value in rows[header_index]]
    return headers, rows[header_index + 1 :]


def _column(headers: list[str], name: str, *aliases: str) -> int:
    normalized = {_clean_header(header): index for index, header in enumerate(headers) if header}
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


def _party(visit_number: Any, patient_name: Any) -> str:
    visit = str(visit_number or "").strip()
    patient = str(patient_name or "").strip()
    return f"{visit}-{patient}" if visit and patient else visit or patient


def _datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, str):
        text = value.strip()
        for pattern in (
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %I:%M %p",
            "%d-%m-%Y %H:%M",
            "%d-%m-%Y %H:%M:%S",
            "%d-%m-%Y %I:%M %p",
            "%d/%m/%y %H:%M",
            "%d/%m/%y %H:%M:%S",
            "%d/%m/%y %I:%M %p",
            "%d-%m-%y %H:%M",
            "%d-%m-%y %H:%M:%S",
            "%d-%m-%y %I:%M %p",
            "%d/%m/%Y",
            "%d/%m/%y",
            "%d-%m-%Y",
            "%d-%m-%y",
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%m/%d/%Y",
        ):
            try:
                return datetime.strptime(text, pattern)
            except ValueError:
                continue
    return None


def _date_text(value: Any, *, four_digit_year: bool) -> str:
    parsed = _datetime(value)
    if parsed is None:
        return str(value or "").strip()
    return parsed.strftime("%d/%m/%Y" if four_digit_year else "%d/%m/%y")


def _date_sort_key(value: Any) -> datetime:
    return _datetime(value) or datetime.min


def _amount(value: Any) -> Any:
    if value in (None, ""):
        return 0
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return value


def _identifier(value: Any) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _bank(mode: Any) -> str:
    text = str(mode or "").strip()
    key = _clean_header(text)
    if key in {"upi", "card", "ubi", "unionbankofindia"}:
        return "UBI"
    if key in {"eft", "indianoverseasbankca3762", "iob"}:
        return "Indian Overseas Bank C.A 3762"
    if key == "cash":
        return "CASH"
    return text


def sales_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    _column(headers, "Bill Number")
    _column(headers, "Bill Date")
    _column(headers, "Visit Number")
    _column(headers, "Patient Name")
    _column(headers, "Services / Charge Name")
    _column(headers, "Item Value")
    _column(headers, "Discount")

    indexed_rows: list[tuple[str, int, list[Any]]] = []
    for index, row in enumerate(source):
        if not _has_values(row):
            continue
        bill = _get(row, headers, "Bill Number")
        bill_date = _get(row, headers, "Bill Date")
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        amount = _amount(_get(row, headers, "Item Value"))
        result = [
            bill, _date_text(bill_date, four_digit_year=True), party, "Auto Sales", None, "Tamil Nadu",
            None, None, None, None, 0, amount, 0, 0, 0, 0, 0, amount,
            party, _amount(_get(row, headers, "Discount")), _get(row, headers, "Services / Charge Name"),
        ]
        indexed_rows.append((str(bill or ""), index, result))
    indexed_rows.sort(key=lambda item: (item[0], item[1]))
    return [row for _, _, row in indexed_rows]


def receipt_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    _column(headers, "Receipt Number")
    _column(headers, "Receipt Date")
    _column(headers, "Visit Number")
    _column(headers, "Patient Name")
    _column(headers, "Mode")
    _column(headers, "Receipt Value")

    indexed_rows: list[tuple[datetime, int, list[Any]]] = []
    for index, row in enumerate(source):
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        receipt_date = _get(row, headers, "Receipt Date")
        output = [
            _date_text(receipt_date, four_digit_year=False),
            _get(row, headers, "Receipt Number"),
            _bank(_get(row, headers, "Mode")),
            party,
            _amount(_get(row, headers, "Receipt Value")),
            _get(row, headers, "DD / Cheque Info"),
            None,
            party,
        ]
        indexed_rows.append((_date_sort_key(receipt_date), index, output))
    indexed_rows.sort(key=lambda item: (item[0], item[1]))
    return [row for _, _, row in indexed_rows]


def refund_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    indexed_rows: list[tuple[datetime, int, list[Any]]] = []
    for index, row in enumerate(source):
        if not _has_values(row):
            continue
        payment_date = _get(row, headers, "Payment Date")
        party = _party(_get(row, headers, "Visit No", "Visit Number"), _get(row, headers, "Patient Name"))
        output = [
            _date_text(payment_date, four_digit_year=True),
            _get(row, headers, "Payment Number"),
            _bank(_get(row, headers, "Mode")),
            party,
            _amount(_get(row, headers, "Payment Value")),
            _get(row, headers, "Cheque No."),
            None,
            party,
        ]
        indexed_rows.append((_date_sort_key(payment_date), index, output))
    indexed_rows.sort(key=lambda item: (item[0], item[1]))
    return [row for _, _, row in indexed_rows]


def paytm_rows(filename: str, payload: bytes) -> list[list[Any]]:
    """Map Card & UPI Payments transactions into the sample Paytm sheet."""
    headers, source = _rows_from_upload(filename, payload)
    date_column = _column(
        headers, "Date", "Transaction Date", "Payment Date", "Receipt Date", "Paid Date", "Created Date"
    )
    receipt_column = _column(
        headers,
        "Rec No",
        "Receipt Number",
        "Receipt No",
        "Receipt ID",
        "Transaction ID",
        "Transaction Reference",
        "Transaction Ref No",
        "Transaction Number",
        "Transaction No",
        "Payment Reference Number",
        "Payment Reference",
        "Payment ID",
        "Ref Number",
        "Ref No",
        "Reference Number",
        "UTR Number",
        "UTR",
        "Order ID",
    )
    account_column = _column(
        headers,
        "Credit Account",
        "Account Name",
        "Ledger",
        "Ledger Name",
        "Charge Item",
        "Services / Charge Name",
        "Item / Service",
        "Service Name",
        "Item Name",
        "Particulars",
        "Description",
    )
    amount_column = _column(
        headers,
        "Amount",
        "Payment Amount",
        "Paid Amount",
        "Transaction Amount",
        "Receipt Value",
        "Net Amount",
        "Credit Amount",
    )

    indexed_rows: list[tuple[datetime, int, list[Any]]] = []
    for index, row in enumerate(source, start=2):
        if not _has_values(row):
            continue
        payment_date = _datetime(row[date_column] if date_column < len(row) else None)
        if payment_date is None:
            raise ValueError(f"{filename}: row {index} has a missing or invalid payment date.")
        party = _party(
            _get(row, headers, "Visit Number", "Visit No", "Visit No."),
            _get(row, headers, "Patient Name", "Patient", "Name"),
        )
        narration = _get(row, headers, "Narration", "Remarks", "Patient / Visit", default=None) or party
        output = [
            payment_date,
            _identifier(row[receipt_column] if receipt_column < len(row) else None),
            "UBI",
            row[account_column] if account_column < len(row) else None,
            _amount(row[amount_column] if amount_column < len(row) else None),
            None,
            None,
            narration,
        ]
        indexed_rows.append((payment_date, index, output))

    indexed_rows.sort(key=lambda item: (item[0], item[1]))
    return [row for _, _, row in indexed_rows]


def discount_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    value_column = _column(headers, "Net Discount Value", "Discount Value")
    result: list[list[Any]] = []
    for row in source:
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        bill = _get(row, headers, "Bill No")
        bill_date = _datetime(_get(row, headers, "Bill Time"))
        value = _amount(row[value_column])
        result.append(["Auto Journal", bill, bill_date, "Discount Allowed", value, None, party])
        result.append(["Auto Journal", bill, bill_date, party, None, value, party])
    return result


def doctor_fee_rows(filename: str, payload: bytes, *, voucher_start: int = 943) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    result: list[list[Any]] = []
    sequence = voucher_start
    for row in source:
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        bill_date = _datetime(_get(row, headers, "Bill Date"))
        value = _amount(_get(row, headers, "Net Value"))
        doctor = _get(row, headers, "Doctor")
        voucher_number = f"DOC/{sequence:04d}"
        sequence += 1
        result.append(["Auto Journal", voucher_number, bill_date, "Professional Charges", value, None, party])
        result.append(["Auto Journal", voucher_number, bill_date, doctor, None, value, party])
    return result


def formula_rows(filename: str, payload: bytes) -> list[list[Any]]:
    """Build the sample's single Formula-tab example row from the IP Bill report."""
    headers, source = _rows_from_upload(filename, payload)
    result: list[list[Any]] = []
    for row in source:
        if not _has_values(row):
            continue
        row_number = len(result) + 2
        category = _get(row, headers, "Bill Category")
        if _clean_header(category) == "ipbill":
            category = "IP"
        result.append([
            None,
            _get(row, headers, "Bill No"),
            category,
            _datetime(_get(row, headers, "Bill Time")),
            _get(row, headers, "MRN"),
            _get(row, headers, "IP No #", "IP No"),
            _get(row, headers, "Patient Name"),
            f'=CONCATENATE(F{row_number},"-",G{row_number})',
            None, None, None, None, None,
        ])
        # The sample Formula tab is a one-row helper/example sheet, not an
        # export of every IP Bill transaction.
        break
    return result


def _copy_row_style(sheet: Any, source_row: int, target_row: int, max_column: int) -> None:
    if target_row == source_row:
        return
    if sheet.row_dimensions[source_row].height is not None:
        sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for column in range(1, max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format


def make_workbook(tables: dict[str, tuple[list[str], Iterable[list[Any]]]]) -> bytes:
    """Fill the sanitized sample layout and return one multi-sheet .xlsx file."""
    if not LAYOUT_TEMPLATE.exists():
        raise FileNotFoundError(f"Output layout template not found: {LAYOUT_TEMPLATE}")

    book = load_workbook(LAYOUT_TEMPLATE)
    try:
        for sheet_name, headers in SHEET_HEADERS.items():
            if sheet_name not in book.sheetnames:
                raise ValueError(f"Output layout is missing the {sheet_name} sheet.")
            sheet = book[sheet_name]
            for cell in list(sheet._cells.values()):
                if cell.row > 1:
                    cell.value = None
                    cell.comment = None
                    cell.hyperlink = None

            if sheet_name not in tables:
                data_rows: list[list[Any]] = []
            else:
                _, iterable_rows = tables[sheet_name]
                data_rows = [list(row) for row in iterable_rows]

            max_column = max(sheet.max_column, len(headers))
            for row_number, values in enumerate(data_rows, start=2):
                if row_number > sheet.max_row or not any(
                    sheet.cell(row_number, column).has_style for column in range(1, max_column + 1)
                ):
                    _copy_row_style(sheet, 2, row_number, max_column)
                for column in range(1, max_column + 1):
                    sheet.cell(row_number, column).value = values[column - 1] if column <= len(values) else None

            if sheet.auto_filter.ref:
                min_col, _, max_col, _ = range_boundaries(sheet.auto_filter.ref)
                last_column = get_column_letter(max_col or len(headers))
                sheet.auto_filter.ref = f"{get_column_letter(min_col or 1)}1:{last_column}{len(data_rows) + 1}"

        output = BytesIO()
        book.save(output)
        return output.getvalue()
    finally:
        book.close()
