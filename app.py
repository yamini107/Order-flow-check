"""
Order Reconciliation App
Compares orders across:
  1. TC Orders sheet        (required) - Order No, Invoice No, Seller SKU, Marketplace
  2. Marketplace sheet      (required) - Order No
  3. WMS sheet              (optional) - client_order_id  (matched against TC Invoice No)
Output: Marketplace Name | Order No | Seller SKU | Comments
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Order Checker", page_icon="📦", layout="wide")

NONE_OPT = "-- Not available --"

# Column-name guesses used to pre-select dropdowns
GUESS = {
    "order": ["order no", "order_no", "order number", "order id", "order_id", "orderid", "order item id", "sub order no"],
    "invoice": ["invoice no", "invoice_no", "invoice number", "invoice id", "invoice"],
    "sku": ["seller sku", "seller_sku", "sku", "sku id", "sku code", "seller sku code"],
    "marketplace": ["marketplace", "marketplace name", "channel", "channel name", "platform"],
    "client_order": ["client_order_id", "client order id", "clientorderid", "client_order_no"],
}


# ---------------------------------------------------------------- helpers
def read_file(uploaded, key):
    """Read CSV / Excel. For Excel with several tabs, let the user pick one."""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, dtype=str, keep_default_na=False)
    xls = pd.ExcelFile(uploaded)
    sheet = xls.sheet_names[0]
    if len(xls.sheet_names) > 1:
        sheet = st.selectbox("Select tab", xls.sheet_names, key=f"tab_{key}")
    return pd.read_excel(xls, sheet_name=sheet, dtype=str, keep_default_na=False)


def normalize(series):
    """Clean keys so VLOOKUP-style matching isn't broken by spaces, case or '.0'."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"^'", "", regex=True)
        .str.upper()
    )


def guess_index(columns, kind, allow_none=False):
    opts = ([NONE_OPT] if allow_none else []) + list(columns)
    lowered = [str(c).strip().lower() for c in opts]
    for g in GUESS[kind]:
        if g in lowered:
            return lowered.index(g)
    return 0


def pick(label, df, kind, key, allow_none=False):
    opts = ([NONE_OPT] if allow_none else []) + list(df.columns)
    return st.selectbox(label, opts, index=guess_index(df.columns, kind, allow_none), key=key)


def join_unique(values):
    vals = [v for v in pd.unique(values) if str(v).strip() not in ("", "NAN", "nan", "None")]
    return ", ".join(map(str, vals))


def to_excel(report, summary):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        report.to_excel(xw, sheet_name="Order Check", index=False)
        report[report["Comments"] != "Order available in all sheets"].to_excel(
            xw, sheet_name="Mismatches Only", index=False
        )
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col) + 2
                ws.column_dimensions[col[0].column_letter].width = min(width, 60)
            ws.freeze_panes = "A2"
    return buf.getvalue()


# ---------------------------------------------------------------- UI
st.title("📦 Order Checker – TC vs Marketplace vs WMS")
st.caption("Upload the sheets, map the columns, and download the comparison report.")

c1, c2, c3 = st.columns(3)
with c1:
    st.subheader("1️⃣ TC Orders sheet")
    tc_file = st.file_uploader("Required", type=["xlsx", "xls", "csv"], key="tc")
with c2:
    st.subheader("2️⃣ Marketplace sheet")
    mp_file = st.file_uploader("Required", type=["xlsx", "xls", "csv"], key="mp")
with c3:
    st.subheader("3️⃣ WMS sheet")
    wms_file = st.file_uploader("Optional", type=["xlsx", "xls", "csv"], key="wms")

if not (tc_file and mp_file):
    st.info("Upload at least the TC Orders sheet and the Marketplace sheet to begin.")
    st.stop()

# ---------------------------------------------------------------- column mapping
st.divider()
st.subheader("🔧 Column mapping")
m1, m2, m3 = st.columns(3)

with m1:
    st.markdown("**TC Orders**")
    tc = read_file(tc_file, "tc")
    tc_order = pick("Order No", tc, "order", "tc_order")
    tc_invoice = pick("Invoice No (used for WMS)", tc, "invoice", "tc_inv", allow_none=True)
    tc_sku = pick("Seller SKU", tc, "sku", "tc_sku", allow_none=True)
    tc_mp = pick("Marketplace Name", tc, "marketplace", "tc_mp", allow_none=True)

with m2:
    st.markdown("**Marketplace**")
    mp = read_file(mp_file, "mp")
    mp_order = pick("Order No", mp, "order", "mp_order")
    mp_sku = pick("Seller SKU", mp, "sku", "mp_sku", allow_none=True)
    mp_mp = pick("Marketplace Name", mp, "marketplace", "mp_mp", allow_none=True)
    default_mp_name = st.text_input(
        "Marketplace name (used when no column is available)", value="Marketplace"
    )

wms = None
with m3:
    st.markdown("**WMS**")
    if wms_file:
        wms = read_file(wms_file, "wms")
        wms_client = pick("client_order_id", wms, "client_order", "wms_client")
        wms_sku = pick("SKU (optional)", wms, "sku", "wms_sku", allow_none=True)
        if tc_invoice == NONE_OPT:
            st.warning("Select the Invoice No column in TC Orders to compare with WMS.")
    else:
        st.caption("No WMS sheet uploaded – WMS check will be skipped.")

use_wms = wms is not None and tc_invoice != NONE_OPT

# ---------------------------------------------------------------- comparison
if st.button("▶️ Run order check", type="primary", use_container_width=True):
    # --- build clean frames
    tcd = pd.DataFrame({
        "key": normalize(tc[tc_order]),
        "order": tc[tc_order].astype(str).str.strip(),
        "invoice": normalize(tc[tc_invoice]) if tc_invoice != NONE_OPT else "",
        "sku": tc[tc_sku].astype(str).str.strip() if tc_sku != NONE_OPT else "",
        "mp": tc[tc_mp].astype(str).str.strip() if tc_mp != NONE_OPT else "",
    })
    tcd = tcd[tcd["key"].ne("") & tcd["key"].ne("NAN")]

    mpd = pd.DataFrame({
        "key": normalize(mp[mp_order]),
        "order": mp[mp_order].astype(str).str.strip(),
        "sku": mp[mp_sku].astype(str).str.strip() if mp_sku != NONE_OPT else "",
        "mp": mp[mp_mp].astype(str).str.strip() if mp_mp != NONE_OPT else "",
    })
    mpd = mpd[mpd["key"].ne("") & mpd["key"].ne("NAN")]

    tc_keys, mp_keys = set(tcd["key"]), set(mpd["key"])
    wms_keys = set()
    if use_wms:
        wmsd = pd.DataFrame({
            "key": normalize(wms[wms_client]),
            "raw": wms[wms_client].astype(str).str.strip(),
            "sku": wms[wms_sku].astype(str).str.strip() if wms_sku != NONE_OPT else "",
        })
        wmsd = wmsd[wmsd["key"].ne("") & wmsd["key"].ne("NAN")]
        wms_keys = set(wmsd["key"])

    # --- group per order (one row per order, SKUs joined)
    tc_g = tcd.groupby("key").agg(
        order=("order", "first"), invoice=("invoice", join_unique),
        sku=("sku", join_unique), mp=("mp", join_unique)).reset_index()
    mp_g = mpd.groupby("key").agg(
        order=("order", "first"), sku=("sku", join_unique), mp=("mp", join_unique)).reset_index()

    merged = tc_g.merge(mp_g, on="key", how="outer", suffixes=("_tc", "_mp"))

    rows = []
    for _, r in merged.iterrows():
        in_tc, in_mp = r["key"] in tc_keys, r["key"] in mp_keys
        comments = []
        if not in_tc:
            comments.append("Order not available in TC Orders sheet")
        if not in_mp:
            comments.append("Order not available in Marketplace sheet")
        if use_wms and in_tc:
            invoices = [i.strip() for i in str(r["invoice"]).split(",") if i.strip()]
            if not invoices:
                comments.append("Invoice No missing in TC – cannot check WMS")
            elif not any(i in wms_keys for i in invoices):
                comments.append(f"Invoice No {', '.join(invoices)} not available in WMS sheet")

        mp_name = (r.get("mp_tc") if isinstance(r.get("mp_tc"), str) and r.get("mp_tc") else None) \
            or (r.get("mp_mp") if isinstance(r.get("mp_mp"), str) and r.get("mp_mp") else None) \
            or default_mp_name
        sku = (r.get("sku_tc") if isinstance(r.get("sku_tc"), str) and r.get("sku_tc") else None) \
            or (r.get("sku_mp") if isinstance(r.get("sku_mp"), str) and r.get("sku_mp") else "")
        order = r["order_tc"] if isinstance(r["order_tc"], str) else r["order_mp"]

        rows.append({
            "Marketplace Name": mp_name,
            "Order No": order,
            "Seller SKU": sku,
            "Comments": "; ".join(comments) if comments else "Order available in all sheets",
        })

    # --- WMS records whose client_order_id doesn't match any TC invoice
    if use_wms:
        tc_invoices = set()
        for inv in tc_g["invoice"]:
            tc_invoices.update(i.strip() for i in str(inv).split(",") if i.strip())
        extra = wmsd[~wmsd["key"].isin(tc_invoices)].groupby("key").agg(
            raw=("raw", "first"), sku=("sku", join_unique)).reset_index()
        for _, r in extra.iterrows():
            rows.append({
                "Marketplace Name": default_mp_name,
                "Order No": r["raw"],
                "Seller SKU": r["sku"],
                "Comments": f"client_order_id {r['raw']} available in WMS but not in TC Orders (Invoice No)",
            })

    report = pd.DataFrame(rows, columns=["Marketplace Name", "Order No", "Seller SKU", "Comments"])
    ok_mask = report["Comments"] == "Order available in all sheets"

    summary = pd.DataFrame({
        "Metric": [
            "Orders in TC Orders sheet", "Orders in Marketplace sheet",
            "Records in WMS sheet" if use_wms else "WMS check",
            "Orders matched in all sheets", "Orders with issues",
        ],
        "Value": [
            len(tc_keys), len(mp_keys),
            len(wms_keys) if use_wms else "Skipped",
            int(ok_mask.sum()), int((~ok_mask).sum()),
        ],
    })
    st.session_state["report"], st.session_state["summary"] = report, summary

# ---------------------------------------------------------------- results
if "report" in st.session_state:
    report, summary = st.session_state["report"], st.session_state["summary"]
    ok_mask = report["Comments"] == "Order available in all sheets"

    st.divider()
    st.subheader("📊 Results")
    k = st.columns(4)
    k[0].metric("TC orders", summary.iloc[0, 1])
    k[1].metric("Marketplace orders", summary.iloc[1, 1])
    k[2].metric("✅ Matched", int(ok_mask.sum()))
    k[3].metric("⚠️ Issues", int((~ok_mask).sum()))

    only_issues = st.toggle("Show only mismatches", value=True)
    search = st.text_input("🔍 Search order no / SKU")
    view = report[~ok_mask] if only_issues else report
    if search:
        s = search.strip().upper()
        view = view[view["Order No"].str.upper().str.contains(s, na=False)
                    | view["Seller SKU"].str.upper().str.contains(s, na=False)]
    st.dataframe(view, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download output sheet (Excel)",
        data=to_excel(report, summary),
        file_name=f"order_check_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
