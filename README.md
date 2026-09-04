# Vehicle Parts Stock - KMN

A Streamlit app for tracking vehicle-parts stock (Issue / Receive / Add
transactions) for two offices — **Chilaw** and **Palavi** — each with its
own login and its own isolated stock records, backed by a shared Google
Sheet.

## Features

- **Per-office login** — Mrs. Hiruni (Chilaw) and Mr. Sampath (Palavi).
  Each officer only ever sees and enters records for their own office.
- **Record Entering** — Event Type, Date, Main Category, Sub Category 1,
  Sub Category 2, Quantity, UOM, GRN NO (for Receive), To/From,
  Description. Category dropdowns are built from data already entered,
  with an "➕ Add new..." option so the taxonomy can grow over time
  without editing code.
- **View Stock** — pick a Main Category and see a ledger table:
  `Date | To/From | Description` on the left, then one column per
  **Sub Category 1** (grouped) → **Sub Category 2**, with the quantity
  for that transaction in the matching cell. The header stays fixed while
  you scroll, and a bold **Balance (as of today)** row stays pinned to
  the bottom, giving current stock per sub-category at a glance.
- Live balance hint while recording an **Issue**, so officers can see
  what's currently in stock for that item before saving.
- CSV export of the filtered ledger.
- Signed quantities: `Receive` and `Add` increase stock, `Issue`
  decreases it.

## How data is stored

Everything is stored in one worksheet, `Stock`, inside a single Google
Sheet, with an `Office` column on every row. The app filters every read
and stamps every write with the logged-in officer's office, so Chilaw and
Palavi data never mix, even though they share the same sheet/app
deployment.

Columns: `Event Type, Date, Main Category, Sub Category 1, Sub Category 2,
Quantity, UOM, GRN NO, To/From, Description, Office, Entered By,
Timestamp` (the last two are added automatically for an audit trail).

## Setup

### 1. Create the Google Sheet + service account

1. Create a new Google Sheet (any name). Copy its ID from the URL:
   `https://docs.google.com/spreadsheets/d/THIS_PART_IS_THE_ID/edit`.
2. In [Google Cloud Console](https://console.cloud.google.com/), create a
   project (or reuse one), then enable **Google Sheets API** and
   **Google Drive API**.
3. Create a **Service Account**, then a JSON key for it — download it.
4. Open the Google Sheet and **Share** it with the service account's
   `client_email` (found in the JSON) as **Editor**.
5. The app creates the `Stock` worksheet and header row automatically the
   first time it runs.

### 2. Configure secrets

Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and
fill in:

- `sheet_id` — the Google Sheet ID from step 1.
- `[gcp_service_account]` — paste every field from the downloaded JSON key.
- `[credentials.hiruni]` / `[credentials.sampath]` — set `name`, `office`,
  and a `password_hash` generated with:

  ```bash
  pip install bcrypt
  python generate_password_hash.py
  ```

  Run it once per officer and paste each printed hash in.

**Never commit `secrets.toml`** — it's already in `.gitignore`.

### 3. Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

### 4. Deploy (Streamlit Community Cloud)

1. Push this folder to a GitHub repo (secrets.toml will be excluded by
   `.gitignore` — that's intentional).
2. On [share.streamlit.io](https://share.streamlit.io), create a new app
   pointing at `app.py` in that repo.
3. In the app's **Settings → Secrets**, paste the full contents of your
   local `.streamlit/secrets.toml`.
4. Deploy. Share the app URL with Mrs. Hiruni and Mr. Sampath.

## Project structure

```
app.py                          # main app: routing between Record Entering / View
lib/auth.py                     # login form, bcrypt password check
lib/sheets.py                   # Google Sheets read/write (gspread)
lib/categories.py                # "select existing or add new" dropdown helper
lib/stock_table.py               # builds the pivoted HTML ledger table for View
generate_password_hash.py       # one-off CLI to create bcrypt hashes
.streamlit/secrets.toml.example # template — copy to secrets.toml and fill in
requirements.txt
```

## Notes / possible next steps

- Currently one shared sheet with an `Office` filter keeps things simple
  to administer; if you'd rather have fully separate spreadsheets per
  office, swap `sheet_id` for an office-keyed lookup in `lib/sheets.py`.
- Add a third "Admin" login (e.g. a head-office role) that can see both
  offices at once, if that's ever needed — the office filter in `app.py`
  is the only place that would need to change.
- GRN NO is only asked for on `Receive` events; adjust in `app.py` if you
  want it available for other event types too.
