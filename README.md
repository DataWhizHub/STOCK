# Vehicle Parts Stock - KMN

A single-file Streamlit app for tracking vehicle-parts stock (Issue /
Receive / Add) for two offices — **Chilaw** and **Palavi** — each with
its own self-service login and isolated records, backed by a shared
Google Sheet.

**Everything is in one file, `app.py`** — no sub-folders.

## Logins (self-service)

The Google Sheet has a `Users` worksheet, auto-created and pre-seeded
with a row for each office but no username/password. The first time an
office opens the app they'll see a collapsed **"🆕 First time here? Set
up your office login"** panel on the login screen — they expand it, pick
their office, and choose their own username and password (bcrypt-hashed
before it's written to the sheet). After that they just log in normally,
and can change their password anytime from **🔑 Change password** in the
sidebar. To reset an office's login yourself, clear that office's
`Username`/`PasswordHash` cells in the `Users` sheet.

## Sections

- **📥 Record Entering** — Event Type, Date, Main Category, Sub Category 1,
  Sub Category 2, Quantity (0 is allowed — e.g. for a zero-stock
  adjustment), UOM, GRN NO (Receive only), To/From, Description.
  Dropdowns grow from existing data, plus "➕ Add new...". **All fields
  reset to blank after a successful save.**

- **📊 View Stock** — pick a Main Category, see the ledger table: Date /
  To-From / Description, then one column per Sub Category 1 (grouped) →
  Sub Category 2, quantity in the matching cell. Sticky header, bold
  **Balance (as of today)** row pinned to the bottom, rest scrolls. The
  **Download this table as CSV** button exports exactly what's shown —
  same columns, same Balance row — not the raw ledger.

- **✏️ Edit Records** — pick any past transaction for your office from a
  dropdown, edit any field and save, or delete it outright.

- **📬 Notifications** — when an office records an **Issue** whose
  To/From is the *other* office's name (e.g. Chilaw issues stock "To"
  Palavi), it automatically creates a pending transfer. The receiving
  office sees it here (with a badge count in the sidebar menu) and taps
  **✅ Mark as Received**, which logs a matching **Receive** entry in
  their own stock automatically. Sent-but-not-yet-received transfers and
  completed transfer history are also shown here.

## Repo structure (flat)

```
app.py
requirements.txt
.streamlit/secrets.toml.example
.gitignore
```

Keep `app.py` at the repo root on GitHub (or set Streamlit Cloud's "Main
file path" to wherever it actually sits).

## Setup

### 1. Google Sheet + service account
1. Create a Google Sheet, copy its ID from the URL.
2. In Google Cloud Console, enable **Google Sheets API** and **Google
   Drive API**, create a **Service Account**, download its JSON key.
3. Share the Sheet with the service account's `client_email` as **Editor**.
4. The app auto-creates the `Stock`, `Users`, and `Transfers` worksheets
   (with `Users` pre-seeded for Chilaw/Palavi) on first run.

### 2. Secrets
Copy `.streamlit/secrets.toml.example` → `.streamlit/secrets.toml`, fill in
`sheet_id` and `[gcp_service_account]`. **Never commit `secrets.toml`**
(already in `.gitignore`).

### 3. Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

### 4. Deploy on Streamlit Community Cloud
1. Push this folder's contents to GitHub — confirm `app.py` is visible
   at the repo root.
2. Create the app with **Main file path = `app.py`**.
3. Settings → Secrets → paste your `secrets.toml` contents.
4. Deploy.

## Notes
- `Receive` and `Add` increase stock, `Issue` decreases it.
- Quantity can be 0 — there's no "must be greater than zero" check.
- GRN NO is only asked for on `Receive`.
- Blank Sub Category 2 is grouped under "General" in the View table.
- Transfer detection matches on the office *name* in To/From (case-
  insensitive) — so an Issue with To/From exactly "Palavi" or "Chilaw"
  triggers the notification; anything else (a supplier name, etc.) is
  treated as a normal external Issue, unchanged.
