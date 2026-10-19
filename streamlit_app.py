"""
Daily Sales vs Payments Reconciliation - web app.
Run with:  streamlit run app.py
"""
import hmac
import os
import tempfile
from datetime import date

import pandas as pd
import streamlit as st

from pipeline import run_reconciliation

st.set_page_config(page_title="Daily Reconciliation", page_icon="🧾", layout="wide")


# ---------- optional password (set APP_PASSWORD env var to turn on) ----------
def _configured_password():
    pw = os.environ.get("APP_PASSWORD")
    if pw:
        return pw
    try:  # Streamlit Cloud keeps secrets here
        return st.secrets.get("APP_PASSWORD")
    except Exception:
        return None


pw = _configured_password()
if os.environ.get("HOSTED") == "1" and not pw:
    st.error("This hosted copy has no APP_PASSWORD configured, so it is locked. "
             "Set APP_PASSWORD in the host's environment/secrets.")
    st.stop()
if pw:
    if st.session_state.get("authed") is not True:
        entered = st.text_input("Password", type="password")
        if entered:
            if hmac.compare_digest(entered.encode(), pw.encode()):
                st.session_state["authed"] = True
                st.rerun()
            st.error("Wrong password.")
        st.stop()

st.title("🧾 Daily Sales vs Payments Reconciliation")
st.caption("Upload the day's files. Every payment source is matched against all shops' bills together, "
           "so cross-shop payments and extra card machines are handled automatically. "
           "Split payments are not flagged - only the digital portion is checked.")


def save(upload, folder):
    path = os.path.join(folder, upload.name)
    with open(path, "wb") as f:
        f.write(upload.getbuffer())
    return path


def table_to_csv(df, folder, name):
    df = df.dropna(subset=["amount"])
    df = df[df["amount"] > 0]
    if df.empty:
        return None
    path = os.path.join(folder, name)
    df.to_csv(path, index=False)
    return path


# ------------------------------- INPUTS --------------------------------
c1, c2 = st.columns([1, 2])
business_date = c1.date_input("Business date", value=date.today())
shop_names = [s.strip() for s in c2.text_input(
    "Shop names (comma separated)", value="Smart Mart, Smart Fashions").split(",") if s.strip()]

st.subheader("1. Billing exports (one per shop)")
pos_uploads = {}
cols = st.columns(max(len(shop_names), 1))
for col, shop in zip(cols, shop_names):
    pos_uploads[shop] = col.file_uploader(f"{shop} - sales Excel", type=["xlsx"], key=f"pos_{shop}")

st.subheader("2. Payment sources")
b1, b2 = st.columns(2)
bank_uploads = b1.file_uploader("Bank statement PDFs (Federal Bank format)", type=["pdf"],
                                accept_multiple_files=True)
phonepe_uploads = b2.file_uploader("PhonePe / EDC settlement CSVs", type=["csv"],
                                   accept_multiple_files=True)

st.markdown("**Card machines** - add as many as you have (any shop). "
            "Upload a CSV, or just type the slip's transactions into the table.")
n_machines = st.number_input("Number of card machines today", min_value=0, max_value=8, value=1, step=1)
machine_inputs = []
EMPTY = pd.DataFrame({"time": pd.Series(dtype="str"), "amount": pd.Series(dtype="float"),
                      "ref": pd.Series(dtype="str")})
for i in range(int(n_machines)):
    with st.expander(f"Card machine {i + 1}", expanded=(i == 0)):
        k1, k2 = st.columns(2)
        label = k1.text_input("Machine name", value=f"Card Machine {i + 1}", key=f"mlabel{i}")
        hint = k2.selectbox("Which shop mainly uses it? (optional hint)", ["Not sure"] + shop_names, key=f"mhint{i}")
        up = st.file_uploader("CSV (columns: time, amount, ref)", type=["csv"], key=f"mcsv{i}")
        edited = st.data_editor(EMPTY, num_rows="dynamic", key=f"medit{i}", width="stretch",
                                column_config={"time": "Time (HH:MM:SS)", "amount": st.column_config.NumberColumn("Amount", format="%.2f"),
                                               "ref": "Auth / ref"})
        machine_inputs.append((label, hint, up, edited))

with st.expander("Direct bank transfers (optional)"):
    t_edit = st.data_editor(EMPTY, num_rows="dynamic", key="transfers", width="stretch")

# --------------------------------- RUN ---------------------------------
if st.button("Run reconciliation", type="primary"):
    missing = [s for s, u in pos_uploads.items() if u is None]
    if missing:
        st.error("Upload the billing export for: " + ", ".join(missing))
    elif not (bank_uploads or phonepe_uploads or machine_inputs):
        st.error("Upload at least one payment source.")
    else:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                pos = {s: save(u, tmp) for s, u in pos_uploads.items()}
                banks = [(f"Bank UPI - {u.name[:25]}" if len(bank_uploads) > 1 else "Federal Bank UPI", save(u, tmp)) for u in bank_uploads]
                pps = [(f"PhonePe" if len(phonepe_uploads) == 1 else f"PhonePe ({j + 1})", save(u, tmp)) for j, u in enumerate(phonepe_uploads)]
                cms = []
                for j, (label, hint, up, edited) in enumerate(machine_inputs):
                    path = save(up, tmp) if up is not None else table_to_csv(edited, tmp, f"machine{j}.csv")
                    if path:
                        cms.append((label, path, None if hint == "Not sure" else hint))
                trs = []
                tp = table_to_csv(t_edit, tmp, "transfers.csv")
                if tp:
                    trs.append(("Direct Transfer", tp))
                out = os.path.join(tempfile.gettempdir(), f"Reconciliation_{business_date}.xlsx")
                summary = run_reconciliation(str(business_date), pos, banks, pps, cms, trs, out, log=lambda x: None)
                with open(out, "rb") as f:
                    st.session_state["xlsx"] = f.read()
                st.session_state["summary"] = summary
        except Exception as e:  # show a readable message instead of a stack trace
            st.session_state.pop("summary", None)
            st.error(f"Could not complete the reconciliation: {e}")

# ------------------------------- RESULTS -------------------------------
if "summary" in st.session_state:
    s = st.session_state["summary"]
    st.divider()
    st.header(f"Results - {s['date']}")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total net sales", f"₹{s['total_net']:,.0f}")
    m2.metric("Digital confirmed", f"₹{s['total_digital']:,.0f}")
    m3.metric("Implied cash (not verified)", f"₹{s['implied_cash']:,.0f}",
              help="Net sales minus digital matched to bills. No cash count is used.")

    st.subheader("By shop")
    st.dataframe(pd.DataFrame(s["shops"]).T.rename(columns={
        "bills": "Bills", "net": "Net sales", "recorded_card": "Till 'Card' total",
        "confirmed_digital": "Digital confirmed", "implied_cash": "Implied cash"}).style.format(
        {"Net sales": "₹{:,.0f}", "Till 'Card' total": "₹{:,.0f}", "Digital confirmed": "₹{:,.0f}", "Implied cash": "₹{:,.0f}"}),
        width="stretch")

    st.subheader("Where each source's money went")
    rows = []
    for name, v in s["sources"].items():
        row = {"Source": name, "Transactions": v["count"], "Total": v["total"]}
        row.update({f"→ {shop}": amt for shop, amt in v["per_shop"].items()})
        rows.append(row)
    st.dataframe(pd.DataFrame(rows).set_index("Source"), width="stretch")
    st.caption("An amount under a shop different from the source's usual shop is cross-shop routing.")

    st.subheader("Items needing a look")
    info = lambda st: st.startswith("Matched - ") and "Review" not in st
    issues = [i for i in s["issues"] if not info(i["status"])]
    cross = [i for i in s["issues"] if info(i["status"])]
    if issues:
        st.dataframe(pd.DataFrame(issues), width="stretch", hide_index=True)
    else:
        st.success("No bills need review.")
    if cross:
        with st.expander(f"{len(cross)} bills paid through another shop's source (money confirmed received)"):
            st.dataframe(pd.DataFrame(cross), width="stretch", hide_index=True)
    if s["unassigned"]:
        st.warning(f"{len(s['unassigned'])} payments (₹{sum(p.amount for p in s['unassigned']):,.0f}) matched no bill:")
        st.dataframe(pd.DataFrame([{"Amount": p.amount, "Source": p.source, "Ref": p.ref, "Time": p.time_str} for p in s["unassigned"]]),
                     width="stretch", hide_index=True)

    st.download_button("⬇️ Download full Excel report", st.session_state["xlsx"],
                       file_name=f"Reconciliation_{s['date']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
