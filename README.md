# Vehicle Parts Stock - KMN

A single-file Streamlit app for tracking vehicle-parts stock (Issue /
Receive / Add) for two offices — **Chilaw** and **Palavi** — each with
its own login and isolated records, backed by a shared Google Sheet.

**Everything is in one file, `app.py`** — no sub-folders, so nothing can
go missing on upload.

## How logins work now

You (the admin) don't need to generate passwords for anyone. The Google
Sheet gets a second worksheet, `Users`, auto-created and pre-seeded with
two rows — one for Chilaw, one for Palavi — with a name but **no
username/password yet**.

The first time either office opens the app, they'll see **"First time
here? Set up your office login"** on the login screen: they pick their
office, enter their name, and choose their own username and password.
It's hashed with bcrypt before it's written to the sheet — nothing is
ever stored in plain text. After that, they just log in normally, and
can change their password anytime from **🔑 Change password** in the
sidebar.

If you ever need to reset an office's login yourself, just clear that
office's `Username` and `PasswordHash` cells in the `Users` worksheet —
the "set up your login" prompt will reappear for that office.

## Features
- Self-service per-office login (see above)
- Record Entering: Event Type, Date, Main Category, Sub Category 1,
  Sub Category 2, Quantity, UOM, GRN NO (Receive only), To/From,
  Description — dropdowns grow from existing data, plus "➕ Add new..."
- View Stock: pick a Main Category, see a ledger table — Date / To-From /
  Description, then one column per Sub Category 1 (grouped) → Sub
  Category 2, quantity in the matching cell. Sticky header, bold
  **Balance (as of today)** row pinned to the bottom, rest scrolls.
- Live stock-balance hint while recording an Issue
- CSV export

## Repo structure (flat)

```
app.py
requirements.txt
.streamlit/secrets.toml.example
.gitignore
```

Make sure `app.py` sits directly at the repo root on GitHub (or set
Streamlit Cloud's "Main file path" to match wherever it actually is).

## Setup

### 1. Google Sheet + service account
1. Create a Google Sheet, copy its ID from the URL
   (`.../spreadsheets/d/THIS_ID/edit`).
2. In Google Cloud Console, enable **Google Sheets API** and
   **Google Drive API**, create a **Service Account**, download its JSON key.
3. Share the Sheet with the service account's `client_email` as **Editor**.
4. The app creates both the `Stock` and `Users` worksheets (with the
   `Users` sheet pre-seeded for Chilaw/Palavi) automatically on first run.

### 2. Secrets
Copy `.streamlit/secrets.toml.example` → `.streamlit/secrets.toml`, fill in:
- `sheet_id`
- `[gcp_service_account]` — every field from the downloaded JSON

**Never commit `secrets.toml`** (already in `.gitignore`).

### 3. Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

### 4. Deploy on Streamlit Community Cloud
1. Push this folder's contents to a GitHub repo — confirm `app.py` is
   visible at the repo root before deploying.
2. On share.streamlit.io, create an app with **Main file path = `app.py`**.
3. Settings → Secrets → paste your local `secrets.toml` contents.
4. Deploy, then send the app link to Mrs. Hiruni and Mr. Sampath — they
   each set up their own login the first time they open it.

## Notes
- `Receive` and `Add` increase stock, `Issue` decreases it.
- GRN NO is only asked for on `Receive`.
- Blank Sub Category 2 is grouped under "General" in the View table.
- Usernames must be unique across both offices.
