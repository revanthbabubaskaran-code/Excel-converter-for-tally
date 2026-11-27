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
SALES_HIDDEN_COLUMNS = ("E", "G", "H", "I", "J", "K", "M", "N", "O", "P", "Q")
VOUCHER_HEADERS = ["Date", "Rec No", "Bank / Cash", "Credit Account", "Amount", "Cheque No", "Chk Date", "Narration"]
PAYMENT_HEADERS = ["Date", "Payment No", "Bank / Cash", "Debit Account", "Amount", "Cheque No", "Chk Date", "Narration"]
JOURNAL_HEADERS = ["VCH Type", "Bill Number", "Bill Date", "Doctor", "Debit Amt", "Item Value", "Narration"]
DISCOUNT_HEADERS = JOURNAL_HEADERS + [""] * 9
DOCTOR_HEADERS = JOURNAL_HEADERS + [""] * 2
PAYTM_HEADERS = VOUCHER_HEADERS + ["", ""]

SHEET_HEADERS = {
    "Auto Sales": SALES_HEADERS,
    "Advances": VOUCHER_HEADERS,
    "Receipt": VOUCHER_HEADERS,
    "Discount Allowed": DISCOUNT_HEADERS,
    "Doctor Fees": DOCTOR_HEADERS,
    "Auto Payment": PAYMENT_HEADERS,
    "Paytm": PAYTM_HEADERS,
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


def _calendar_date_sort_key(value: Any) -> datetime:
    parsed = _datetime(value)
    return parsed.replace(hour=0, minute=0, second=0, microsecond=0) if parsed else datetime.min


def _alphanumeric_sort_key(value: Any) -> tuple[tuple[int, Any], ...]:
    parts = re.split(r"(\d+)", str(value or "").strip())
    return tuple((1, int(part)) if part.isdigit() else (0, part.casefold()) for part in parts if part)


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
    bill_type_column = _column(headers, "Bill Type")
    _column(headers, "Bill Number")
    _column(headers, "Bill Date")
    _column(headers, "Visit Number")
    _column(headers, "Patient Name")
    _column(headers, "Services / Charge Name")
    _column(headers, "Item Value")

    indexed_rows: list[tuple[datetime, tuple[tuple[int, Any], ...], int, list[Any]]] = []
    for index, row in enumerate(source):
        if not _has_values(row):
            continue
        bill_type = row[bill_type_column] if bill_type_column < len(row) else None
        if str(bill_type or "").strip().casefold() == "op":
            continue
        bill = _get(row, headers, "Bill Number")
        bill_date = _get(row, headers, "Bill Date")
        parsed_date = _datetime(bill_date)
        if parsed_date is None:
            raise ValueError(f"{filename}: row {index + 2} has a missing or invalid bill date.")
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        amount = _amount(_get(row, headers, "Item Value"))
        result = [
            bill, _date_text(bill_date, four_digit_year=True), party, "Auto Sales", None, "Tamil Nadu",
            None, None, None, None, 0, amount, 0, 0, 0, 0, 0, amount,
            party, 0, _get(row, headers, "Services / Charge Name"),
        ]
        indexed_rows.append((_calendar_date_sort_key(bill_date), _alphanumeric_sort_key(bill), index, result))
    indexed_rows.sort(key=lambda item: item[:3])
    return [row for _, _, _, row in indexed_rows]


def receipt_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    _column(headers, "Receipt Number")
    _column(headers, "Receipt Date")
    _column(headers, "Visit Number")
    _column(headers, "Patient Name")
    _column(headers, "Mode")
    _column(headers, "Receipt Value")

    indexed_rows: list[tuple[datetime, tuple[tuple[int, Any], ...], int, list[Any]]] = []
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
        indexed_rows.append((
            _calendar_date_sort_key(receipt_date),
            _alphanumeric_sort_key(_get(row, headers, "Receipt Number")),
            index,
            output,
        ))
    indexed_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    return [row for _, _, _, row in indexed_rows]


def refund_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    indexed_rows: list[tuple[datetime, tuple[tuple[int, Any], ...], int, list[Any]]] = []
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
        indexed_rows.append((
            _calendar_date_sort_key(payment_date),
            _alphanumeric_sort_key(_get(row, headers, "Payment Number")),
            index,
            output,
        ))
    indexed_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    return [row for _, _, _, row in indexed_rows]


def paytm_rows(filename: str, payload: bytes, *, first_rec_no: str) -> list[list[Any]]:
    """Map the Card & UPI statement into the sample Paytm sheet."""
    headers, source = _rows_from_upload(filename, payload)
    date_column = _column(headers, "BANK CREDIT DATE")
    sequence_column = _column(headers, "SL.NO", "SL NO")
    account_column = _column(headers, "DEPART")
    card_upi_column = _column(headers, "CARD/UPI", "CARD / UPI")
    amount_column = _column(headers, "TOTAL")
    name_column = _column(headers, "NAME")
    visit_column = _column(headers, "DEP / IP NO", "DEP /  IP NO", "DEP/IP NO")

    rec_text = str(first_rec_no or "").strip()
    rec_match = re.fullmatch(r"([^\d]*)(\d+)", rec_text)
    if rec_match is None:
        raise ValueError("Enter the first Paytm Rec No with a numeric ending, for example PAYTM01720.")
    rec_prefix, rec_digits = rec_match.groups()
    sequence = int(rec_digits)
    number_width = len(rec_digits)

    candidates: list[tuple[datetime, tuple[tuple[int, Any], ...], int, list[Any]]] = []
    for index, row in enumerate(source, start=2):
        if not _has_values(row):
            continue
        card_upi = row[card_upi_column] if card_upi_column < len(row) else None
        if str(card_upi or "").strip().casefold() == "total":
            continue
        account = row[account_column] if account_column < len(row) else None
        if str(account or "").strip().casefold() in {"advance", "discharge"}:
            continue
        raw_date = row[date_column] if date_column < len(row) else None
        # The statement has footer/summary lines with totals but no bank date.
        if raw_date in (None, ""):
            continue
        payment_date = _datetime(raw_date)
        if payment_date is None:
            raise ValueError(f"{filename}: row {index} has a missing or invalid payment date.")
        name = row[name_column] if name_column < len(row) else None
        visit = row[visit_column] if visit_column < len(row) else None
        visit_text = "" if visit is None else str(visit).strip()
        name_text = "" if name is None else str(name).strip()
        narration = f"{visit_text}-{name_text}"
        source_sequence = row[sequence_column] if sequence_column < len(row) else None
        candidates.append((payment_date, _alphanumeric_sort_key(source_sequence), index, [
            payment_date,
            None,
            "UBI",
            account,
            _amount(row[amount_column] if amount_column < len(row) else None),
            None,
            None,
            narration,
        ]))
    candidates.sort(key=lambda item: (_calendar_date_sort_key(item[0]), item[1], item[2]))
    result: list[list[Any]] = []
    for _, _, _, output in candidates:
        output[1] = f"{rec_prefix}{sequence:0{number_width}d}"
        result.append(output)
        sequence += 1
    return result


def discount_rows(filename: str, payload: bytes) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    value_column = _column(headers, "Net Discount Value", "Discount Value")
    indexed_rows: list[tuple[datetime, tuple[tuple[int, Any], ...], int, list[Any]]] = []
    for index, row in enumerate(source):
        if not _has_values(row):
            continue
        party = _party(_get(row, headers, "Visit Number"), _get(row, headers, "Patient Name"))
        bill = _get(row, headers, "Bill No")
        bill_date = _datetime(_get(row, headers, "Bill Time"))
        if bill_date is not None:
            bill_date = bill_date.replace(hour=0, minute=0, second=0, microsecond=0)
        value = _amount(row[value_column])
        indexed_rows.append((
            _calendar_date_sort_key(bill_date), _alphanumeric_sort_key(bill), index,
            ["Auto Journal", bill, bill_date, "Discount Allowed", value, None, party],
        ))
    indexed_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    result: list[list[Any]] = []
    for _, _, _, debit_row in indexed_rows:
        credit_row = ["Auto Journal", debit_row[1], debit_row[2], debit_row[6], None, debit_row[4], debit_row[6]]
        debit_row[5] = None
        result.extend((debit_row, credit_row))
    return result


def doctor_fee_rows(
    filename: str,
    payload: bytes,
    *,
    voucher_start: str | int = "DOC/0943",
) -> list[list[Any]]:
    headers, source = _rows_from_upload(filename, payload)
    voucher_text = str(voucher_start).strip()
    voucher_match = re.fullmatch(r"([^\d]*)(\d+)", voucher_text)
    if voucher_match is None:
        raise ValueError("Enter the first Bill Number with a numeric ending, for example DOC/0943.")
    voucher_prefix, voucher_digits = voucher_match.groups()
    sequence = int(voucher_digits)
    number_width = len(voucher_digits)
    result: list[list[Any]] = []
    dated_source = [
        (
            _calendar_date_sort_key(_get(row, headers, "Bill Date")),
            _alphanumeric_sort_key(_get(row, headers, "Bill Number")),
            index,
            row,
        )
        for index, row in enumerate(source)
        if _has_values(row)
    ]
    dated_source.sort(key=lambda item: (item[0], item[1], item[2]))
    for _, _, _, row in dated_source:
        visit = _get(row, headers, "Visit Number")
        party = _party(visit, _get(row, headers, "Patient Name"))
        bill_date = _datetime(_get(row, headers, "Bill Date"))
        value = _amount(_get(row, headers, "Net Value"))
        doctor = _get(row, headers, "Doctor")
        voucher_number = f"{voucher_prefix}{sequence:0{number_width}d}"
        sequence += 1
        result.append(["Auto Journal", voucher_number, bill_date, "Professional Charges", value, None, party])
        result.append(["Auto Journal", voucher_number, bill_date, doctor, None, value, party])
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
            if sheet_name == "Auto Sales":
                for column in SALES_HIDDEN_COLUMNS:
                    sheet.column_dimensions[column].hidden = True
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

            if sheet_name == "Discount Allowed" and len(data_rows) >= 2:
                # The supplied Tally sample includes this extra ledger value
                # on the second journal line in column P.
                sheet.cell(3, 16).value = "Discount Allowed"

            if sheet.auto_filter.ref:
                min_col, _, max_col, _ = range_boundaries(sheet.auto_filter.ref)
                last_column = get_column_letter(max_col or len(headers))
                sheet.auto_filter.ref = f"{get_column_letter(min_col or 1)}1:{last_column}{len(data_rows) + 1}"

        output = BytesIO()
        book.save(output)
        return output.getvalue()
    finally:
        book.close()
