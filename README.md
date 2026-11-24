# Excel for Tally

A local Windows desktop app that converts the supplied hospital `.xls` reports into a workbook with the sample output sheet names and columns. It opens as a desktop window and does not use a browser.

The first run asks you to verify an email address with a randomly generated four-digit numeric code, then create a local profile with a 4–8 character password. The code expires after 10 minutes; a newly requested code replaces the previous one. Email IDs are matched without regard to capitalization. Later runs require those credentials. Passwords are stored as salted hashes; if **Remember me** is selected, a separate copy is encrypted with Windows DPAPI for autofill on the same Windows account.

Email verification uses SMTP. Before first use, set these environment variables on the Windows account that runs the app: `EXCEL_TALLY_SMTP_HOST`, `EXCEL_TALLY_SMTP_PORT`, `EXCEL_TALLY_SMTP_USERNAME`, `EXCEL_TALLY_SMTP_PASSWORD`, and `EXCEL_TALLY_SMTP_FROM`. Port 465 uses SSL; other configured ports use STARTTLS. For Gmail, use an app password rather than your normal account password. If SMTP is not configured or email delivery fails, profile setup remains blocked and the app does not reveal the code.

## Run on Windows

Install Python 3.14 or newer and [uv](https://docs.astral.sh/uv/). Then double-click `Excel for tally.bat`. The launcher installs the project dependencies and opens the app window.

You can also open PowerShell in this folder and run:

```powershell
uv sync
uv run tally-final-upload-build
```

## Reports to select

Select these six Excel reports in the app:

- Charge Item Report
- Doctor Fees Report
- IP Advance Report
- IP Discount Report
- IP Receipt Report
- IP Refund Report

The IP Bill Report is not used because the supplied sample output has no matching sheet. The PDF bank statement is ignored.

## Output notes

The workbook contains `Auto Sales`, `Advances`, `Receipt`, `Discount Allowed`, `Doctor Fees`, `Auto Payment`, `Paytm`, and `Formula` sheets. The Paytm and Formula sheets are empty apart from their headers: the sample's 662 Paytm rows are not present in the supplied Excel reports. No bank statement data is read.

The conversion rules are based on the July 2026 examples. Review the generated workbook before importing it into accounting software, especially if report columns or formats change.
