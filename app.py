"""
Vehicle Parts Stock - KMN
Single-file Streamlit app.

Both stock records AND office login credentials live in the same Google
Sheet:
  - "Stock" worksheet: the transaction ledger (Office column isolates data)
  - "Users" worksheet: one row per office (Chilaw, Palavi), pre-seeded
    with the office name but no username/password. The first time an
    office opens the app, they set their own username + password
    (hashed with bcrypt before it's ever written to the sheet). They can
    change their password later from the sidebar.
"""

import html as html_lib
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

STOCK_HEADERS = [
    "Event Type", "Date", "Main Category", "Sub Category 1", "Sub Category 2",
    "Quantity", "UOM", "GRN NO", "To/From", "Description",
    "Office", "Entered By", "Timestamp",
]
USER_HEADERS = ["Office", "Name", "Username", "PasswordHash", "UpdatedAt"]

# The two offices this app serves. "Name" here is just the default shown
# until the officer sets up their profile / changes it themselves.
OFFICE_SEED = [
    {"Office": "Chilaw", "Name": "Mrs. Hiruni"},
    {"Office": "Palavi", "Name": "Mr. Sampath"},
]

SIGN_MAP = {"Issue": -1, "Receive": 1, "Add": 1}
ADD_NEW = "➕ Add new..."
PLACEHOLDER = "-- Select --"


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
    # Seed the two office rows if they aren't there yet
    existing_offices = {r.get("Office", "") for r in ws.get_all_records()}
    for seed in OFFICE_SEED:
        if seed["Office"] not in existing_offices:
            ws.append_row([seed["Office"], seed["Name"], "", "", ""])
    return ws


# =========================================================
# STOCK DATA
# =========================================================
@st.cache_data(ttl=30, show_spinner=False)
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


def append_stock_entry(row: dict) -> None:
    ws = _get_stock_ws()
    ws.append_row([row.get(h, "") for h in STOCK_HEADERS], value_input_option="USER_ENTERED")
    load_stock.clear()


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
    if cell is None:
        ws.append_row([office, name, username, password_hash, datetime.now().isoformat(timespec="seconds")])
    else:
        row = cell.row
        ws.update(f"A{row}:E{row}", [[office, name, username, password_hash, datetime.now().isoformat(timespec="seconds")]])
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
    st.info("👋 One or both offices haven't set up a login yet. Set yours up below.")
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
            with st.expander("🆕 First time here? Set up your office login", expanded=True):
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
# UI HELPERS (record entering / view)
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


def build_stock_table_html(data: pd.DataFrame, as_of: pd.Timestamp) -> str:
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


# =========================================================
# PAGES
# =========================================================
def render_entry(df_office, user, office):
    st.title("📥 Record Entering")
    st.caption(f"New stock transaction — {office} office")

    c1, c2 = st.columns(2)
    with c1:
        event_type = st.selectbox("Event Type *", ["Issue", "Receive", "Add"])
    with c2:
        entry_date = st.date_input("Date *", value=date.today())

    main_cat = selectbox_with_add("Main Category", df_office["Main Category"].tolist(), "main_cat")

    sub1_options = df_office.loc[df_office["Main Category"] == main_cat, "Sub Category 1"].tolist() if main_cat else []
    sub1 = selectbox_with_add("Sub Category 1", sub1_options, "sub1")

    sub2_options = df_office.loc[
        (df_office["Main Category"] == main_cat) & (df_office["Sub Category 1"] == sub1), "Sub Category 2"
    ].tolist() if sub1 else []
    sub2 = selectbox_with_add("Sub Category 2", sub2_options, "sub2", required=False)

    c3, c4 = st.columns(2)
    with c3:
        quantity = st.number_input("Quantity *", min_value=0.0, step=1.0, format="%.2f")
    with c4:
        uom = selectbox_with_add("UOM", df_office["UOM"].tolist(), "uom", required=False)

    grn_no = ""
    if event_type == "Receive":
        grn_no = st.text_input("GRN NO")

    to_from = st.text_input("To / From *", placeholder="Supplier (Receive) / Recipient (Issue) / Source (Add)")
    description = st.text_area("Description", height=80)

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
        if quantity <= 0:
            errors.append("Quantity (must be greater than 0)")
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
        st.success("✅ Saved successfully!")
        st.balloons()
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
    st.markdown(build_stock_table_html(data, today), unsafe_allow_html=True)

    with st.expander("⬇️ Export raw data for this Main Category"):
        cols = ["Date", "Event Type", "Sub Category 1", "Sub Category 2",
                "Quantity", "UOM", "GRN NO", "To/From", "Description", "Entered By"]
        st.dataframe(data.sort_values("Date")[cols], use_container_width=True, hide_index=True)
        st.download_button(
            "Download CSV",
            data.sort_values("Date").to_csv(index=False).encode("utf-8"),
            file_name=f"{office}_{main_cat}_stock.csv",
            mime="text/csv",
        )


# =========================================================
# MAIN
# =========================================================
def main():
    user = require_login()
    office = user["office"]

    st.sidebar.markdown(f"### 👤 {user['name']}")
    st.sidebar.markdown(f"**Office:** {office}")
    st.sidebar.divider()
    page = st.sidebar.radio("Menu", ["📥 Record Entering", "📊 View Stock"], label_visibility="collapsed")
    st.sidebar.divider()
    change_password_ui(user)
    logout_button()

    df_all = load_stock()
    df_office = df_all[df_all["Office"] == office].copy() if not df_all.empty else df_all

    if page.startswith("📥"):
        render_entry(df_office, user, office)
    else:
        render_view(df_office, office)


if __name__ == "__main__":
    main()
