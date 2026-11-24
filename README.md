# Excel for Tally

A local Windows desktop app that converts the supplied hospital `.xls` reports into a workbook with the sample output sheet names and columns. It opens as a desktop window and does not use a browser.

## Device licensing

Each Windows installation needs an owner-approved device license. The app submits its request code to the online service. The owner reviews the request in a protected dashboard and clicks **Approve**; the service then signs and returns the license. The user clicks **Check status** in the app to activate it. Licenses are bound to the Windows MachineGuid, so copying a licensed app to another device requires a new request. Reinstalling Windows can require a new license. Users with the source code can alter the app's checks.

### Free Cloudflare deployment

The service is designed for Cloudflare Workers and D1. The free tier currently includes 100,000 Worker requests per day; D1 includes 5 million rows read per day, 100,000 rows written per day, and 5 GB total storage. See [Workers limits](https://developers.cloudflare.com/workers/platform/limits/) and [D1 pricing](https://developers.cloudflare.com/d1/platform/pricing/).

1. Create a Cloudflare account and install Node.js. In `licensing_service`, run `npm install` and `npx wrangler login`.
2. Create the D1 database with `npx wrangler d1 create excel-for-tally-licenses`. Put the returned database ID into `licensing_service/wrangler.jsonc` in place of `REPLACE_WITH_D1_DATABASE_ID`.
3. Apply the schema with `npx wrangler d1 migrations apply LICENSE_DB --remote`.
4. Set the owner dashboard token as a Cloudflare secret named `LICENSE_ADMIN_TOKEN`. Generate a long random value and keep it private. In PowerShell, run `npx wrangler secret put LICENSE_ADMIN_TOKEN`, then paste the value when prompted.
5. Set `LICENSE_SIGNING_KEY` as a Cloudflare secret. On the owner Windows account, run `uv run python tools/license_service_key.py export-signing-key` and copy its one-line output. From `licensing_service`, run `npx wrangler secret put LICENSE_SIGNING_KEY` and paste the output directly at the prompt. Never put this key in Git, chat, or the app. It matches the public verification key in `src/tally_final_upload_build/license_public_key.txt`.
6. Deploy with `npx wrangler deploy`. Cloudflare prints the service URL. Enter that HTTPS URL in the app's license screen; it is saved for that Windows account. Open `<service-url>/admin`, enter the owner dashboard token, and approve pending device requests there. The license is created by the service only after you click **Approve**.

The private signing key stays in Cloudflare after deployment. The local `.license_signing_key.dpapi` file is only used to transfer that key once and remains ignored by Git. The service URL is stored locally for the current Windows user. Cloudflare account login and deployment are required to make the service live.

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
