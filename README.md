# Excel for Tally

A local Windows desktop app that converts the supplied hospital `.xls` reports into a workbook with the sample output sheet names and columns. It opens as a desktop window and does not use a browser.

## Device licensing

Each Windows installation needs an owner-approved device license. The app submits its request code to the online service. The owner reviews the request in a protected dashboard and clicks **Approve**; the service then creates a unique signed license for that device. The user clicks **Check status** in the app to activate it. Licenses are bound to the Windows MachineGuid, so copying a license to another device will not activate it and requires a new request. The same device can check again to recover its approved license. Reinstalling Windows can require a new license. Users with the source code can alter the app's checks.

### Free Cloudflare deployment

The service is designed for Cloudflare Workers and D1. The free tier currently includes 100,000 Worker requests per day; D1 includes 5 million rows read per day, 100,000 rows written per day, and 5 GB total storage. See [Workers limits](https://developers.cloudflare.com/workers/platform/limits/) and [D1 pricing](https://developers.cloudflare.com/d1/platform/pricing/).

1. Create a Cloudflare account and install Node.js. In `licensing_service`, run `npm install` and `npx wrangler login`.
2. Create the D1 database with `npx wrangler d1 create excel-for-tally-licenses`. Put the returned database ID into `licensing_service/wrangler.jsonc` in place of `REPLACE_WITH_D1_DATABASE_ID`.
3. Apply the schema with `npx wrangler d1 migrations apply LICENSE_DB --remote`.
4. Create an owner dashboard token once with `uv run python tools/license_service_key.py init-admin-token`. It is encrypted with Windows DPAPI and ignored by Git. Add it to Cloudflare as a secret named `LICENSE_ADMIN_TOKEN`. To copy it without displaying it in the terminal, run `uv run python tools/license_service_key.py export-admin-token | Set-Clipboard`, then paste it into the Cloudflare secret form.
5. Set `LICENSE_SIGNING_KEY` as a Cloudflare secret. On the owner Windows account, run `uv run python tools/license_service_key.py export-signing-key | Set-Clipboard`, then paste the value into the Cloudflare secret form. Never put this key in Git, chat, or the app. It matches the public verification key in `src/tally_final_upload_build/license_public_key.txt`.
6. Deploy with `npx wrangler deploy`. Cloudflare prints the service URL. Enter that HTTPS URL in the app's license screen; it is saved for that Windows account. Open `<service-url>/admin`, enter the owner dashboard token, and approve pending device requests there. The license is created by the service only after you click **Approve**.

The deployed service is at `https://excel-for-tally-licensing.revtechsolutions-tn.workers.dev`. The app uses this service by default unless another service URL is saved for the current Windows user. The Worker and D1 database are deployed; the owner dashboard and license issuance require the `LICENSE_ADMIN_TOKEN` and `LICENSE_SIGNING_KEY` Cloudflare secrets to be configured. The private signing key stays in Cloudflare after setup. The local `.license_signing_key.dpapi` file remains ignored by Git.

The first run lets you create a local profile with an email ID and a 4–8 character password; email verification is not required. Email IDs are matched without regard to capitalization. Later runs require those credentials. Passwords are stored as salted hashes; if **Remember me** is selected, a separate copy is encrypted with Windows DPAPI for autofill on the same Windows account.

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
