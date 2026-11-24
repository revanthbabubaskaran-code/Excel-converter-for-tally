# Excel for Tally

A local Windows desktop app that converts the supplied hospital `.xls` reports into a workbook with the sample output sheet names and columns. It opens as a desktop window and does not use a browser.

## Device licensing

Each Windows installation needs its own owner-issued license. The app displays a device request code; send it to the software owner, then paste the returned license into the app. A license is signed for that Windows installation and will not activate on a different device. This uses the Windows MachineGuid as its device identifier, so reinstalling Windows can require a new license.

The owner initializes the signing key once on the license-issuing Windows account with `uv run python tools/issue_license.py init`. Keep `.license_signing_key.dpapi` on that computer and never share or commit it. The tool protects it with Windows DPAPI. Commit/distribute `src/tally_final_upload_build/license_public_key.txt`; it contains only the public verification key. To issue a license, run `uv run python tools/issue_license.py issue EFT1-<request-code>` and send the output token to the requester. The license check is intended for distributed app copies; users with the source code can alter the app.

The first run asks you to verify an email address with a randomly generated four-digit numeric code, then create a local profile with a 4–8 character password. The code expires after 10 minutes; a newly requested code replaces the previous one. Email IDs are matched without regard to capitalization. Later runs require those credentials. Passwords are stored as salted hashes; if **Remember me** is selected, a separate copy is encrypted with Windows DPAPI for autofill on the same Windows account.

To send profile verification codes, click **Connect Gmail account** in setup. The app opens the user's default browser for Google sign-in and consent, then sends through the Gmail API; it does not ask for a Gmail password or show the connected email address. Gmail refresh tokens are protected with Windows DPAPI for the current Windows account.

One-time Google setup is required: create a Google Cloud project, enable the Gmail API, configure the OAuth consent screen, and create an OAuth client of type **Desktop app**. The app asks for this public client ID the first time **Connect Gmail account** is clicked and saves it locally. The only requested OAuth scope is `https://www.googleapis.com/auth/gmail.send`. During testing, add the signing-in Gmail account as a test user if Google requires it. See Google's [Desktop app OAuth guide](https://developers.google.com/identity/protocols/oauth2/native-app) and [Gmail API send guide](https://developers.google.com/workspace/gmail/api/guides/sending). If OAuth is not configured or email delivery fails, profile creation remains blocked and the app does not reveal the verification code.

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
