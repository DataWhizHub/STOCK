"""
Vehicle Parts Stock - KMN
Single-file Streamlit app.

Google Sheet worksheets used:
  - "Stock"     : the transaction ledger (Office column isolates data)
  - "Users"     : one row per office, holds each office's own username/password
  - "Transfers" : inter-office stock transfers awaiting acknowledgement
                  (created automatically when an office Issues stock
                  "To/From" the other office; the other office sees it
                  under Notifications and can mark it Received, which
                  auto-adds a matching Receive entry to their own stock)
"""

import html as html_lib
import uuid
from datetime import date, datetime

import bcrypt
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Vehicle Parts Stock - KMN", page_icon="🔧", layout="wide")

# =========================================================
# CONFIG
# =========================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]
STOCK_SHEET = "Stock"
USERS_SHEET = "Users"
TRANSFERS_SHEET = "Transfers"

STOCK_HEADERS = [
    "Event Type", "Date", "Main Category", "Sub Category 1", "Sub Category 2",
    "Quantity", "UOM", "GRN NO", "To/From", "Description",
    "Office", "Entered By", "Timestamp",
]
USER_HEADERS = ["Office", "Name", "Username", "PasswordHash", "UpdatedAt"]
TRANSFER_HEADERS = [
    "TransferID", "From Office", "To Office", "Date", "Main Category",
    "Sub Category 1", "Sub Category 2", "Quantity", "UOM", "GRN NO", "Description",
    "Status", "Issued By", "Issued At", "Received By", "Received At",
]

OFFICES = ["Chilaw", "Palavi"]
OFFICE_SEED = [
    {"Office": "Chilaw", "Name": "Mrs. Hiruni"},
    {"Office": "Palavi", "Name": "Mr. Sampath"},
]

SIGN_MAP = {"Issue": -1, "Receive": 1, "Add": 1}
ADD_NEW = "➕ Add new..."
PLACEHOLDER = "-- Select --"


def other_office(office: str) -> str:
    return "Palavi" if office == "Chilaw" else "Chilaw"


# =========================================================
# GOOGLE SHEETS - connection
# =========================================================
@st.cache_resource(show_spinner=False)
def _get_client() -> gspread.Client:
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def _get_spreadsheet():
    return _get_client().open_by_key(st.secrets["sheet_id"])


@st.cache_resource(show_spinner=False)
def _get_stock_ws():
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(STOCK_SHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=STOCK_SHEET, rows=2000, cols=len(STOCK_HEADERS) + 2)
        ws.append_row(STOCK_HEADERS)
    if not ws.row_values(1):
        ws.append_row(STOCK_HEADERS)
    return ws


@st.cache_resource(show_spinner=False)
def _get_users_ws():
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(USERS_SHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=USERS_SHEET, rows=20, cols=len(USER_HEADERS) + 1)
        ws.append_row(USER_HEADERS)
    if not ws.row_values(1):
        ws.append_row(USER_HEADERS)
    existing_offices = {r.get("Office", "") for r in ws.get_all_records()}
    for seed in OFFICE_SEED:
        if seed["Office"] not in existing_offices:
            ws.append_row([seed["Office"], seed["Name"], "", "", ""])
    return ws


@st.cache_resource(show_spinner=False)
def _get_transfers_ws():
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(TRANSFERS_SHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TRANSFERS_SHEET, rows=2000, cols=len(TRANSFER_HEADERS) + 2)
        ws.append_row(TRANSFER_HEADERS)
    if not ws.row_values(1):
        ws.append_row(TRANSFER_HEADERS)
    return ws


def _records_with_row_numbers(ws, headers) -> pd.DataFrame:
    """Like get_all_records(), but keeps the actual sheet row number in
    a `_row` column so single rows can be updated/deleted later."""
    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return pd.DataFrame(columns=headers + ["_row"])
    data_rows = all_values[1:]
    records = []
    for i, r in enumerate(data_rows, start=2):
        rec = {h: (r[idx] if idx < len(r) else "") for idx, h in enumerate(headers)}
        rec["_row"] = i
        records.append(rec)
    return pd.DataFrame(records)


# =========================================================
# STOCK DATA
# =========================================================
@st.cache_data(ttl=20, show_spinner=False)
def load_stock() -> pd.DataFrame:
    ws = _get_stock_ws()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=STOCK_HEADERS)

    for col in STOCK_HEADERS:
        if col not in df.columns:
            df[col] = ""

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0.0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    text_cols = ["Sub Category 1", "Sub Category 2", "UOM", "To/From", "Description",
                 "Main Category", "GRN NO", "Office", "Entered By", "Event Type"]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df[STOCK_HEADERS]


@st.cache_data(ttl=15, show_spinner=False)
def load_stock_with_rows() -> pd.DataFrame:
    return _records_with_row_numbers(_get_stock_ws(), STOCK_HEADERS)


def _clear_stock_caches():
    load_stock.clear()
    load_stock_with_rows.clear()


def append_stock_entry(row: dict) -> None:
    ws = _get_stock_ws()
    ws.append_row([row.get(h, "") for h in STOCK_HEADERS], value_input_option="USER_ENTERED")
    _clear_stock_caches()


def update_stock_row(row_number: int, row: dict) -> None:
    ws = _get_stock_ws()
    values = [row.get(h, "") for h in STOCK_HEADERS]
    last_col = chr(ord("A") + len(STOCK_HEADERS) - 1)  # 13 headers -> "M"
    ws.update(f"A{row_number}:{last_col}{row_number}", [values])
    _clear_stock_caches()


def delete_stock_row(row_number: int) -> None:
    ws = _get_stock_ws()
    ws.delete_rows(row_number)
    _clear_stock_caches()


# =========================================================
# USERS / AUTH
# =========================================================
@st.cache_data(ttl=10, show_spinner=False)
def load_users() -> pd.DataFrame:
    ws = _get_users_ws()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        df = pd.DataFrame(columns=USER_HEADERS)
    for col in USER_HEADERS:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)
    return df


def _hash_pw(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def _check_pw(password: str, password_hash: str) -> bool:
    if not password_hash:
        return False
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except ValueError:
        return False


def save_user_credentials(office: str, name: str, username: str, password_hash: str) -> None:
    ws = _get_users_ws()
    cell = ws.find(office, in_column=1)
    row_values = [office, name, username, password_hash, datetime.now().isoformat(timespec="seconds")]
    if cell is None:
        ws.append_row(row_values)
    else:
        ws.update(f"A{cell.row}:E{cell.row}", [row_values])
    load_users.clear()


def authenticate(username: str, password: str):
    users = load_users()
    username = username.strip().lower()
    match = users[users["Username"].str.lower() == username]
    if match.empty:
        return None
    row = match.iloc[0]
    if _check_pw(password, row["PasswordHash"]):
        return {"office": row["Office"], "name": row["Name"], "username": username}
    return None


def offices_without_login(users: pd.DataFrame) -> list:
    return users.loc[users["Username"].str.strip() == "", "Office"].tolist()


# =========================================================
# LOGIN / PROFILE SETUP UI
# =========================================================
def profile_setup_form(users: pd.DataFrame, pending_offices: list):
    with st.form("profile_setup"):
        office = st.selectbox("Your office", pending_offices)
        default_name = users.loc[users["Office"] == office, "Name"].values
        name = st.text_input("Your name", value=default_name[0] if len(default_name) else "")
        username = st.text_input("Choose a username")
        pw1 = st.text_input("Choose a password", type="password")
        pw2 = st.text_input("Confirm password", type="password")
        submitted = st.form_submit_button("Create my login", type="primary", use_container_width=True)

    if submitted:
        errors = []
        if not name.strip():
            errors.append("Name")
        if not username.strip():
            errors.append("Username")
        if len(pw1) < 4:
            errors.append("Password must be at least 4 characters")
        if pw1 != pw2:
            errors.append("Passwords do not match")
        existing_usernames = users["Username"].str.lower().tolist()
        if username.strip().lower() in existing_usernames:
            errors.append("That username is already taken")

        if errors:
            st.error(" · ".join(errors))
            return

        save_user_credentials(office, name.strip(), username.strip().lower(), _hash_pw(pw1))
        st.success(f"Login created for {office}! You can now log in below.")
        st.rerun()


def login_form():
    st.markdown("<h2 style='text-align:center;'>🔧 Vehicle Parts Stock - KMN</h2>", unsafe_allow_html=True)
    users = load_users()
    pending = offices_without_login(users)

    _, col2, _ = st.columns([1, 1.2, 1])
    with col2:
        if pending:
            with st.expander("🆕 First time here? Set up your office login", expanded=False):
                profile_setup_form(users, pending)
            st.divider()

        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", type="primary", use_container_width=True)
        if submitted:
            user = authenticate(username, password)
            if user:
                st.session_state["user"] = user
                st.rerun()
            else:
                st.error("Invalid username or password.")


def require_login():
    if "user" not in st.session_state:
        login_form()
        st.stop()
    return st.session_state["user"]


def logout_button():
    if st.sidebar.button("🚪 Log out", use_container_width=True):
        st.session_state.pop("user", None)
        st.rerun()


def change_password_ui(user):
    with st.sidebar.expander("🔑 Change password"):
        with st.form("change_pw_form"):
            current = st.text_input("Current password", type="password")
            new1 = st.text_input("New password", type="password")
            new2 = st.text_input("Confirm new password", type="password")
            ok = st.form_submit_button("Update password", use_container_width=True)
        if ok:
            check = authenticate(user["username"], current)
            if not check:
                st.error("Current password is incorrect.")
            elif len(new1) < 4:
                st.error("New password must be at least 4 characters.")
            elif new1 != new2:
                st.error("New passwords do not match.")
            else:
                save_user_credentials(user["office"], user["name"], user["username"], _hash_pw(new1))
                st.success("Password updated.")


# =========================================================
# TRANSFERS (inter-office notifications)
# =========================================================
@st.cache_data(ttl=10, show_spinner=False)
def load_transfers() -> pd.DataFrame:
    return _records_with_row_numbers(_get_transfers_ws(), TRANSFER_HEADERS)


def _clear_transfer_caches():
    load_transfers.clear()


def create_transfer(row: dict) -> None:
    ws = _get_transfers_ws()
    ws.append_row([row.get(h, "") for h in TRANSFER_HEADERS], value_input_option="USER_ENTERED")
    _clear_transfer_caches()


def mark_transfer_received(transfer_row: dict, received_by: str) -> None:
    ws = _get_transfers_ws()
    row_number = int(transfer_row["_row"])
    updated = dict(transfer_row)
    updated["Status"] = "Received"
    updated["Received By"] = received_by
    updated["Received At"] = datetime.now().isoformat(timespec="seconds")
    values = [updated.get(h, "") for h in TRANSFER_HEADERS]
    last_col = chr(ord("A") + len(TRANSFER_HEADERS) - 1)  # 16 headers -> "P"
    ws.update(f"A{row_number}:{last_col}{row_number}", [values])
    _clear_transfer_caches()

    desc = f"Transfer from {transfer_row['From Office']}"
    if transfer_row.get("Description"):
        desc += f" — {transfer_row['Description']}"

    append_stock_entry({
        "Event Type": "Receive",
        "Date": date.today().isoformat(),
        "Main Category": transfer_row["Main Category"],
        "Sub Category 1": transfer_row["Sub Category 1"],
        "Sub Category 2": transfer_row["Sub Category 2"],
        "Quantity": transfer_row["Quantity"],
        "UOM": transfer_row["UOM"],
        "GRN NO": transfer_row.get("GRN NO", ""),
        "To/From": transfer_row["From Office"],
        "Description": desc,
        "Office": transfer_row["To Office"],
        "Entered By": received_by,
        "Timestamp": datetime.now().isoformat(timespec="seconds"),
    })


def pending_incoming_transfers(office: str) -> pd.DataFrame:
    df = load_transfers()
    if df.empty:
        return df
    return df[(df["To Office"] == office) & (df["Status"] == "Pending")]


# =========================================================
# SHARED UI HELPERS
# =========================================================
def selectbox_with_add(label: str, options: list, key: str, required: bool = True) -> str:
    options = sorted({o for o in options if o})
    choice_list = [PLACEHOLDER] + options + [ADD_NEW]
    label_display = f"{label} *" if required else label
    choice = st.selectbox(label_display, choice_list, key=f"{key}_choice")
    if choice == ADD_NEW:
        return st.text_input(f"New {label}", key=f"{key}_new").strip()
    if choice == PLACEHOLDER:
        return ""
    return choice


def _esc(v) -> str:
    return html_lib.escape(str(v)) if v not in (None, "") else ""


def _fmt_num(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    v = float(v)
    return f"{v:.0f}" if v == int(v) else f"{v:g}"


def compute_pivot_and_balance(data: pd.DataFrame, as_of: pd.Timestamp):
    data = data.copy()
    data.loc[data["Sub Category 2"] == "", "Sub Category 2"] = "General"
    data["Signed Qty"] = data["Quantity"] * data["Event Type"].map(SIGN_MAP).fillna(1)

    pivot = data.pivot_table(
        index=["Date", "To/From", "Description"],
        columns=["Sub Category 1", "Sub Category 2"],
        values="Signed Qty",
        aggfunc="sum",
    )
    pivot = pivot.sort_index(level="Date")

    balance_src = data[data["Date"] <= as_of]
    balance = balance_src.groupby(["Sub Category 1", "Sub Category 2"])["Signed Qty"].sum()

    sub1_order = list(dict.fromkeys(c[0] for c in pivot.columns))
    sub2_by_sub1 = {c1: [] for c1 in sub1_order}
    for c1, c2 in pivot.columns:
        sub2_by_sub1[c1].append(c2)

    return pivot, balance, sub1_order, sub2_by_sub1


def stock_table_html(pivot, balance, sub1_order, sub2_by_sub1, as_of) -> str:
    css = """
<style>
.spk-wrap { max-height: 560px; overflow-y: auto; border: 1px solid rgba(128,128,128,.4); border-radius: 8px; }
table.spk-table { border-collapse: collapse; width: 100%; font-size: 13px; }
table.spk-table th, table.spk-table td {
    border: 1px solid rgba(128,128,128,.35); padding: 6px 10px; text-align: center; white-space: nowrap;
}
table.spk-table thead th { position: sticky; background: #262730; color: #fafafa; z-index: 3; }
table.spk-table thead tr:nth-child(1) th { top: 0; }
table.spk-table thead tr:nth-child(2) th { top: 35px; }
table.spk-table td:nth-child(-n+3), table.spk-table th:nth-child(-n+3) { text-align: left; }
table.spk-table tfoot td { position: sticky; bottom: 0; background: #143d14; color: #fafafa; font-weight: 700; z-index: 3; }
</style>
"""
    thead = "<thead><tr>"
    thead += '<th rowspan="2">Date</th><th rowspan="2">To/From</th><th rowspan="2">Description</th>'
    for c1 in sub1_order:
        thead += f'<th colspan="{len(sub2_by_sub1[c1])}">{_esc(c1)}</th>'
    thead += "</tr><tr>"
    for c1 in sub1_order:
        for c2 in sub2_by_sub1[c1]:
            thead += f"<th>{_esc(c2)}</th>"
    thead += "</tr></thead>"

    tbody = "<tbody>"
    if pivot.empty:
        colspan = 3 + sum(len(v) for v in sub2_by_sub1.values()) or 4
        tbody += f'<tr><td colspan="{colspan}" style="text-align:center;">No records for this Main Category yet.</td></tr>'
    else:
        for (d, tf, desc), row in pivot.iterrows():
            d_str = d.strftime("%Y-%m-%d") if pd.notna(d) else ""
            tbody += f"<tr><td>{_esc(d_str)}</td><td>{_esc(tf)}</td><td>{_esc(desc)}</td>"
            for c1 in sub1_order:
                for c2 in sub2_by_sub1[c1]:
                    tbody += f"<td>{_fmt_num(row.get((c1, c2)))}</td>"
            tbody += "</tr>"
    tbody += "</tbody>"

    tfoot = f'<tfoot><tr><td colspan="3">Balance (as of {as_of.strftime("%Y-%m-%d")})</td>'
    for c1 in sub1_order:
        for c2 in sub2_by_sub1[c1]:
            tfoot += f"<td>{_fmt_num(balance.get((c1, c2), 0))}</td>"
    tfoot += "</tr></tfoot>"

    return f'{css}<div class="spk-wrap"><table class="spk-table">{thead}{tbody}{tfoot}</table></div>'


def stock_table_export_df(pivot, balance, sub1_order, sub2_by_sub1, as_of) -> pd.DataFrame:
    """Same rows/columns as the displayed table (including the Balance
    row), flattened for a CSV download."""
    flat_cols = [f"{c1} - {c2}" for c1 in sub1_order for c2 in sub2_by_sub1[c1]]
    rows = []
    for (d, tf, desc), row in pivot.iterrows():
        d_str = d.strftime("%Y-%m-%d") if pd.notna(d) else ""
        rec = {"Date": d_str, "To/From": tf, "Description": desc}
        for c1 in sub1_order:
            for c2 in sub2_by_sub1[c1]:
                val = row.get((c1, c2))
                rec[f"{c1} - {c2}"] = "" if pd.isna(val) else val
        rows.append(rec)

    balance_rec = {"Date": "", "To/From": "", "Description": f"Balance (as of {as_of.strftime('%Y-%m-%d')})"}
    for c1 in sub1_order:
        for c2 in sub2_by_sub1[c1]:
            balance_rec[f"{c1} - {c2}"] = balance.get((c1, c2), 0)
    rows.append(balance_rec)

    return pd.DataFrame(rows, columns=["Date", "To/From", "Description"] + flat_cols)


# =========================================================
# PAGES
# =========================================================
ENTRY_KEYS = [
    "re_event_type", "re_date", "re_main_cat_choice", "re_main_cat_new",
    "re_sub1_choice", "re_sub1_new", "re_sub2_choice", "re_sub2_new",
    "re_qty", "re_uom_choice", "re_uom_new", "re_grn", "re_to_from", "re_desc",
]


def render_entry(df_office, user, office):
    st.title("📥 Record Entering")
    st.caption(f"New stock transaction — {office} office")

    c1, c2 = st.columns(2)
    with c1:
        event_type = st.selectbox("Event Type *", ["Issue", "Receive", "Add"], key="re_event_type")
    with c2:
        entry_date = st.date_input("Date *", value=date.today(), key="re_date")

    main_cat = selectbox_with_add("Main Category", df_office["Main Category"].tolist(), "re_main_cat")

    sub1_options = df_office.loc[df_office["Main Category"] == main_cat, "Sub Category 1"].tolist() if main_cat else []
    sub1 = selectbox_with_add("Sub Category 1", sub1_options, "re_sub1")

    sub2_options = df_office.loc[
        (df_office["Main Category"] == main_cat) & (df_office["Sub Category 1"] == sub1), "Sub Category 2"
    ].tolist() if sub1 else []
    sub2 = selectbox_with_add("Sub Category 2", sub2_options, "re_sub2", required=False)

    c3, c4 = st.columns(2)
    with c3:
        quantity = st.number_input("Quantity *", min_value=0.0, step=1.0, format="%.2f", key="re_qty")
    with c4:
        uom = selectbox_with_add("UOM", df_office["UOM"].tolist(), "re_uom", required=False)

    grn_no = ""
    if event_type == "Receive":
        grn_no = st.text_input("GRN NO", key="re_grn")

    to_from = st.text_input(
        "To / From *",
        placeholder=f"e.g. a supplier, or '{other_office(office)}' for an inter-office transfer",
        key="re_to_from",
    )
    description = st.text_area("Description", height=80, key="re_desc")

    if event_type == "Issue" and main_cat and sub1:
        sign = df_office["Event Type"].map(SIGN_MAP).fillna(1)
        mask = (df_office["Main Category"] == main_cat) & (df_office["Sub Category 1"] == sub1)
        mask &= df_office["Sub Category 2"] == sub2 if sub2 else df_office["Sub Category 2"] == ""
        current_balance = (df_office.loc[mask, "Quantity"] * sign[mask]).sum()
        st.info(f"Current balance for **{sub1}{' / ' + sub2 if sub2 else ''}**: **{current_balance:g}**")

    st.write("")
    if st.button("💾 Save", type="primary", use_container_width=True):
        errors = []
        if not main_cat:
            errors.append("Main Category")
        if not sub1:
            errors.append("Sub Category 1")
        if not to_from.strip():
            errors.append("To/From")
        if errors:
            st.error("Please fix the following: " + ", ".join(errors))
            return

        append_stock_entry({
            "Event Type": event_type,
            "Date": entry_date.isoformat(),
            "Main Category": main_cat,
            "Sub Category 1": sub1,
            "Sub Category 2": sub2,
            "Quantity": quantity,
            "UOM": uom,
            "GRN NO": grn_no,
            "To/From": to_from.strip(),
            "Description": description.strip(),
            "Office": office,
            "Entered By": user["name"],
            "Timestamp": datetime.now().isoformat(timespec="seconds"),
        })

        transfer_created = False
        if event_type == "Issue" and to_from.strip().lower() == other_office(office).lower():
            create_transfer({
                "TransferID": uuid.uuid4().hex[:10],
                "From Office": office,
                "To Office": other_office(office),
                "Date": entry_date.isoformat(),
                "Main Category": main_cat,
                "Sub Category 1": sub1,
                "Sub Category 2": sub2,
                "Quantity": quantity,
                "UOM": uom,
                "GRN NO": grn_no,
                "Description": description.strip(),
                "Status": "Pending",
                "Issued By": user["name"],
                "Issued At": datetime.now().isoformat(timespec="seconds"),
                "Received By": "",
                "Received At": "",
            })
            transfer_created = True

        for k in ENTRY_KEYS:
            st.session_state.pop(k, None)

        if transfer_created:
            st.success(f"✅ Saved! {other_office(office)} has been notified to receive this transfer.")
        else:
            st.success("✅ Saved successfully!")
        st.rerun()


def render_view(df_office, office):
    st.title("📊 View Stock")
    st.caption(f"{office} office")

    main_cats = sorted({m for m in df_office["Main Category"].tolist() if m})
    if not main_cats:
        st.info("No records yet for this office.")
        return

    main_cat = st.selectbox("Main Category", main_cats)
    data = df_office[df_office["Main Category"] == main_cat].copy()

    today = pd.Timestamp(date.today())
    pivot, balance, sub1_order, sub2_by_sub1 = compute_pivot_and_balance(data, today)

    st.markdown(stock_table_html(pivot, balance, sub1_order, sub2_by_sub1, today), unsafe_allow_html=True)

    export_df = stock_table_export_df(pivot, balance, sub1_order, sub2_by_sub1, today)
    st.download_button(
        "⬇️ Download this table as CSV",
        export_df.to_csv(index=False).encode("utf-8"),
        file_name=f"{office}_{main_cat}_stock_table.csv",
        mime="text/csv",
    )


def render_edit(office):
    st.title("✏️ Edit Records")
    st.caption(f"{office} office — update or delete a saved transaction")

    df = load_stock_with_rows()
    df_office = df[df["Office"] == office].copy() if not df.empty else df
    if df_office.empty:
        st.info("No records yet for this office.")
        return

    df_office = df_office.sort_values("Date", ascending=False)

    def label_for(r):
        sub = r["Sub Category 1"] + (f" / {r['Sub Category 2']}" if r["Sub Category 2"] else "")
        return f"{r['Date']} · {r['Event Type']} · {r['Main Category']} / {sub} · Qty {r['Quantity']} · {r['To/From']}"

    labels = [label_for(r) for _, r in df_office.iterrows()]
    rows_by_label = {label_for(r): r for _, r in df_office.iterrows()}

    selected_label = st.selectbox("Select a record to edit", labels, key="edit_select")
    rec = rows_by_label[selected_label]

    with st.form("edit_form"):
        c1, c2 = st.columns(2)
        with c1:
            options = ["Issue", "Receive", "Add"]
            idx = options.index(rec["Event Type"]) if rec["Event Type"] in options else 0
            event_type = st.selectbox("Event Type", options, index=idx)
        with c2:
            try:
                default_date = datetime.strptime(rec["Date"], "%Y-%m-%d").date()
            except (ValueError, TypeError):
                default_date = date.today()
            entry_date = st.date_input("Date", value=default_date)

        main_cat = st.text_input("Main Category", value=rec["Main Category"])
        sub1 = st.text_input("Sub Category 1", value=rec["Sub Category 1"])
        sub2 = st.text_input("Sub Category 2", value=rec["Sub Category 2"])

        c3, c4 = st.columns(2)
        with c3:
            try:
                default_qty = float(rec["Quantity"] or 0)
            except ValueError:
                default_qty = 0.0
            quantity = st.number_input("Quantity", min_value=0.0, step=1.0, format="%.2f", value=default_qty)
        with c4:
            uom = st.text_input("UOM", value=rec["UOM"])

        grn_no = st.text_input("GRN NO", value=rec["GRN NO"])
        to_from = st.text_input("To/From", value=rec["To/From"])
        description = st.text_area("Description", value=rec["Description"])

        c5, c6 = st.columns(2)
        with c5:
            save = st.form_submit_button("💾 Save changes", type="primary", use_container_width=True)
        with c6:
            delete = st.form_submit_button("🗑️ Delete record", use_container_width=True)

    if save:
        update_stock_row(int(rec["_row"]), {
            "Event Type": event_type,
            "Date": entry_date.isoformat(),
            "Main Category": main_cat.strip(),
            "Sub Category 1": sub1.strip(),
            "Sub Category 2": sub2.strip(),
            "Quantity": quantity,
            "UOM": uom.strip(),
            "GRN NO": grn_no.strip(),
            "To/From": to_from.strip(),
            "Description": description.strip(),
            "Office": office,
            "Entered By": rec["Entered By"],
            "Timestamp": rec["Timestamp"],
        })
        st.success("Record updated.")
        st.rerun()

    if delete:
        delete_stock_row(int(rec["_row"]))
        st.success("Record deleted.")
        st.rerun()


def render_notifications(office, user):
    st.title("📬 Notifications")
    st.caption(f"Incoming stock transfers for {office} office")

    incoming = pending_incoming_transfers(office)
    if incoming.empty:
        st.info("No pending transfers right now.")
    else:
        for _, t in incoming.iterrows():
            with st.container(border=True):
                sub = t["Sub Category 1"] + (f" / {t['Sub Category 2']}" if t["Sub Category 2"] else "")
                st.markdown(f"**From {t['From Office']}** — {t['Main Category']} / {sub}")
                st.write(f"Quantity: **{t['Quantity']} {t['UOM']}**  ·  Date: {t['Date']}  ·  Issued by: {t['Issued By']}")
                if t.get("Description"):
                    st.caption(t["Description"])
                if st.button("✅ Mark as Received", key=f"recv_{t['_row']}", type="primary"):
                    mark_transfer_received(t.to_dict(), user["name"])
                    st.success("Added to your stock as a Receive entry.")
                    st.rerun()

    all_transfers = load_transfers()
    if not all_transfers.empty:
        sent = all_transfers[(all_transfers["From Office"] == office) & (all_transfers["Status"] == "Pending")]
        if not sent.empty:
            with st.expander(f"⏳ Sent, awaiting {other_office(office)} to receive ({len(sent)})"):
                st.dataframe(
                    sent[["Date", "Main Category", "Sub Category 1", "Sub Category 2", "Quantity", "UOM"]],
                    use_container_width=True, hide_index=True,
                )
        history = all_transfers[
            ((all_transfers["From Office"] == office) | (all_transfers["To Office"] == office))
            & (all_transfers["Status"] == "Received")
        ]
        if not history.empty:
            with st.expander(f"✅ Completed transfers ({len(history)})"):
                st.dataframe(
                    history[["Date", "From Office", "To Office", "Main Category", "Sub Category 1",
                              "Sub Category 2", "Quantity", "UOM", "Received By", "Received At"]]
                    .sort_values("Received At", ascending=False),
                    use_container_width=True, hide_index=True,
                )


# =========================================================
# MAIN
# =========================================================
def main():
    user = require_login()
    office = user["office"]

    pending_count = len(pending_incoming_transfers(office))
    notif_label = f"📬 Notifications ({pending_count})" if pending_count else "📬 Notifications"

    st.sidebar.markdown(f"### 👤 {user['name']}")
    st.sidebar.markdown(f"**Office:** {office}")
    st.sidebar.divider()
    page = st.sidebar.radio(
        "Menu",
        ["📥 Record Entering", "📊 View Stock", "✏️ Edit Records", notif_label],
        label_visibility="collapsed",
    )
    st.sidebar.divider()
    change_password_ui(user)
    logout_button()

    df_all = load_stock()
    df_office = df_all[df_all["Office"] == office].copy() if not df_all.empty else df_all

    if page.startswith("📥"):
        render_entry(df_office, user, office)
    elif page.startswith("📊"):
        render_view(df_office, office)
    elif page.startswith("✏️"):
        render_edit(office)
    else:
        render_notifications(office, user)


if __name__ == "__main__":
    main()
