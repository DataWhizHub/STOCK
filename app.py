import streamlit as st
import pandas as pd
from datetime import date

from lib.auth import require_login, logout_button
from lib.sheets import load_data, append_entry, now_iso
from lib.categories import selectbox_with_add
from lib.stock_table import build_stock_table_html

st.set_page_config(page_title="Vehicle Parts Stock - KMN", page_icon="🔧", layout="wide")

# ---------- Login ----------
user = require_login()
office = user["office"]

# ---------- Sidebar ----------
st.sidebar.markdown(f"### 👤 {user['name']}")
st.sidebar.markdown(f"**Office:** {office}")
st.sidebar.divider()
page = st.sidebar.radio("Menu", ["📥 Record Entering", "📊 View Stock"], label_visibility="collapsed")
st.sidebar.divider()
logout_button()

# ---------- Data (this office only) ----------
df_all = load_data()
df_office = df_all[df_all["Office"] == office].copy() if not df_all.empty else df_all


# =========================================================
# RECORD ENTERING
# =========================================================
def render_entry():
    st.title("📥 Record Entering")
    st.caption(f"New stock transaction — {office} office")

    c1, c2 = st.columns(2)
    with c1:
        event_type = st.selectbox("Event Type *", ["Issue", "Receive", "Add"])
    with c2:
        entry_date = st.date_input("Date *", value=date.today())

    main_cat = selectbox_with_add(
        "Main Category", df_office["Main Category"].tolist(), "main_cat"
    )

    sub1_options = df_office.loc[
        df_office["Main Category"] == main_cat, "Sub Category 1"
    ].tolist() if main_cat else []
    sub1 = selectbox_with_add("Sub Category 1", sub1_options, "sub1")

    sub2_options = df_office.loc[
        (df_office["Main Category"] == main_cat) & (df_office["Sub Category 1"] == sub1),
        "Sub Category 2",
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

    to_from = st.text_input(
        "To / From *",
        placeholder="Supplier name (Receive) / Recipient (Issue) / Source (Add)",
    )
    description = st.text_area("Description", height=80)

    # live balance hint for Issue, so the officer doesn't oversell stock
    if event_type == "Issue" and main_cat and sub1:
        sub2_key = sub2 if sub2 else "General"
        sign = df_office["Event Type"].map({"Issue": -1, "Receive": 1, "Add": 1}).fillna(1)
        mask = (df_office["Main Category"] == main_cat) & (df_office["Sub Category 1"] == sub1)
        if sub2:
            mask &= df_office["Sub Category 2"] == sub2
        else:
            mask &= df_office["Sub Category 2"] == ""
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

        row = {
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
            "Timestamp": now_iso(),
        }
        append_entry(row)
        st.success("✅ Saved successfully!")
        st.balloons()
        st.rerun()


# =========================================================
# VIEW STOCK
# =========================================================
def render_view():
    st.title("📊 View Stock")
    st.caption(f"{office} office")

    main_cats = sorted({m for m in df_office["Main Category"].tolist() if m})
    if not main_cats:
        st.info("No records yet for this office.")
        return

    main_cat = st.selectbox("Main Category", main_cats)
    data = df_office[df_office["Main Category"] == main_cat].copy()

    today = pd.Timestamp(date.today())
    html_table = build_stock_table_html(data, today)
    st.markdown(html_table, unsafe_allow_html=True)

    with st.expander("⬇️ Export raw data for this Main Category"):
        st.dataframe(
            data.sort_values("Date")[
                ["Date", "Event Type", "Sub Category 1", "Sub Category 2",
                 "Quantity", "UOM", "GRN NO", "To/From", "Description", "Entered By"]
            ],
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "Download CSV",
            data.sort_values("Date").to_csv(index=False).encode("utf-8"),
            file_name=f"{office}_{main_cat}_stock.csv",
            mime="text/csv",
        )


# ---------- Route ----------
if page.startswith("📥"):
    render_entry()
else:
    render_view()
