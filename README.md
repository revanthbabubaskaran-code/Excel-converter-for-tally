# Excel for Tally

A local Windows desktop app that converts the supplied hospital `.xls` reports into a workbook with the sample output sheet names and columns. It opens as a desktop window and does not use a browser.

## Device licensing

Each Windows installation needs an owner-approved device license. Setup sends a request to the online service and checks its status automatically while setup remains open. The owner reviews requests in a protected dashboard and clicks **Approve**; the service then creates a signed license for that device. Licenses are bound to the Windows MachineGuid, so another device needs its own approval. Setup also requires a network connection to the license service before installation.

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

Download and run [Excel for Tally Setup.exe](release/Excel%20for%20Tally%20Setup.exe). Setup installs `Excel for Tally.exe` together with its bundled `_internal` runtime folder and dependencies; it does not require a separate Python installation. License requests and checks run automatically in the background. A network connection is required during installation.

To rebuild the Windows packages with Python 3.14 and PyInstaller installed, run these commands from the repository root:

```powershell
pyinstaller --noconfirm --distpath dist/app --workpath build/app "packaging/Excel for Tally.spec"
pyinstaller --noconfirm --distpath dist/setup --workpath build/setup "packaging/Excel for Tally Setup.spec"
```

The setup spec bundles the complete `dist/app/Excel for Tally` directory into the installer. The installed application directory contains the executable and all files from that bundle.

For development from source, install Python 3.14 or newer and [uv](https://docs.astral.sh/uv/), then run `uv sync` and `uv run tally-final-upload-build`.

You can also open PowerShell in this folder and run:

```powershell
uv sync
uv run tally-final-upload-build
```

## Reports to select

Select one or more of these Excel reports in the app:

- Charge Item Report
- Doctor Fees Report
- IP Advance Report
- IP Discount Report
- IP Receipt Report
- IP Refund Report
- Card & UPI Payments

The PDF bank statement is ignored. The output contains seven sheets; sheets whose source report was not selected remain blank.

## Output notes

Choose one `.xlsx` output file. It contains these seven sheets in order: `Auto Sales`, `Advances`, `Receipt`, `Discount Allowed`, `Doctor Fees`, `Auto Payment`, and `Paytm`, following the sample's sheet layout. The `Formula` sheet has been removed. For Paytm, `BANK CREDIT DATE` maps to Date, Rec No increments from the value entered at generation time, Bank / Cash is `UBI`, `DEPART` maps to Credit Account, `TOTAL` maps to Amount, and Narration is `DEP / IP NO-NAME`. Rows with `DEPART` equal to `ADVANCE` or `DISCHARGE` and rows without a bank credit date are skipped. Reports that were not selected leave their corresponding sheets blank. The PDF bank statement is ignored.

The conversion rules are based on the July 2026 examples. Auto Sales skips rows whose `Bill Type` is `OP`, then sorts the rest by Bill Date ascending and alphanumeric Bill Number ascending for matching dates; dates are written as `DD/MM/YYYY`. Columns E and G–J are blank; K, M–Q, and T contain zero. The sample layout hides columns E, G–K, and M–Q. Review the generated workbook before importing it into accounting software, especially if report columns or formats change.
