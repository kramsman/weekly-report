# Service Account Setup for weekly-report

This replaces the OAuth token flow (token.json + credentials.json) with a service account.
No browser popup is ever needed. The JSON key file does not expire.

---

## Step 1 — Create the service account in Google Cloud Console

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and select the project:
   - Account: `tech@centerforcommonground.org`
   - Project name: **weekly-report**
   - Project ID: `weekly-report-490922`
   - Project number: `623535509991`

2. In the left menu: **IAM & Admin → Service Accounts**.

3. Click **+ CREATE SERVICE ACCOUNT**.
   - Name: `weekly-report` (or similar)
   - Description: "Uploads weekly reports to Google Drive"
   - Click **CREATE AND CONTINUE**

4. Skip the optional role/permissions steps (permissions are handled by Drive folder sharing, not IAM).
   Click **DONE**.

5. Click on the new service account in the list to open it.

6. Go to the **KEYS** tab → **ADD KEY → Create new key → JSON → CREATE**.

7. A `.json` file downloads automatically. This is the key file — treat it like a password.

---

## Step 2 — Place the key file

Save the downloaded JSON key file to:

```
~/.config/weekly-report/service_account.json
```

On macOS this expands to `/Users/Denise/.config/weekly-report/service_account.json`.

Make sure this path is in `.gitignore` (already covered by `credentials*.json` — but add explicitly):

```
service_account*.json
```

---

## Step 3 — Share the Drive folders with the service account

The service account has its own Google identity (an email like
`weekly-report@voterletters-reports.iam.gserviceaccount.com`). It cannot access Drive folders
unless they are explicitly shared with it.

**Find the service account email:**
In Google Cloud Console → IAM & Admin → Service Accounts, copy the email shown.

**Share both target folders (Editor access):**

1. Open Google Drive as `tech@centerforcommonground.org`.
2. Navigate to the **VoterLetters Admin Reports** folder
   (ID: `1PU8hcYfE3Vlh5v8Cq60Gup_8lfaFff4b`).
3. Right-click → **Share** → paste the service account email → set role to **Editor** → Send.
4. Repeat for the **VoterLetters Organizer Reports** folder
   (ID: `1BcvXzwyEKGiiWSt27NLaassLE0DMgioP`).

---

## Step 4 — Reinstall bekgoogle after the new code is pushed

After the new `get_serviceaccount_creds.py` and `create_google_services_serviceaccount.py`
modules are committed and pushed to the bekgoogle GitHub repo, reinstall in this project:

```bash
uv sync --upgrade-package bekgoogle
```

---

## Verification

Run the app:

```bash
uv run weekly-report
```

- It should start **without any browser popup**.
- Choose **Upload** and confirm a file uploads to the correct Drive folder.
- Confirm share notification emails are sent to recipients.
- The old `token.json` and `credentials.json` files are no longer needed.

---

## Notes

- Drive notification emails (the "you've been given access" messages) come from
  `drive-shares-noreply@google.com` — the service account can trigger these just like a user can.
- The service account **cannot** send Gmail. If you ever need to send email directly (not via
  the Drive share notification), that would require domain-wide delegation or a different approach.
- The key file never expires, but you can rotate it in Cloud Console if needed.
