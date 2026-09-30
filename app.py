"""
Order Reconciliation App
Compares orders across:
  1. TC Orders sheet        (required) - order_number, invoice_number, custom_sku, channel
  2. Marketplace sheet      (required) - Lazada / Shopee / TikTok export
  3. WMS sheet              (optional) - client_order_id  (matched against TC invoice_number)
Output: Marketplace Name | Order No | Seller SKU | Comments
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Order Checker", page_icon="📦", layout="wide")

NONE_OPT = "-- Not available --"
MARKETPLACES = ["Lazada", "Shopee", "TikTok"]

# ------------------------------------------------------------------------------
# Auto-mapping presets. The first name found in the uploaded sheet is selected.
# Add more column names here if your exports use different headers.
# ------------------------------------------------------------------------------
TC_PRESET = {
    "order":   ["order_number", "order no", "order_no", "order id", "order_id"],
    "invoice": ["invoice_number", "invoice no", "invoice_no", "invoice"],
    "sku":     ["custom_sku", "seller sku", "seller_sku", "sku"],
    "channel": ["channel", "marketplace", "platform"],
}

MP_PRESET = {
    "Lazada": {
        "order": ["orderNumber", "order number", "order_number", "order no"],
        "sku":   ["sellerSku", "seller sku", "seller_sku"],
    },
    "Shopee": {
        "order": ["Order ID", "order_sn", "order sn", "order_id", "order no"],
        "sku":   ["SKU Reference No.", "Parent SKU Reference No.", "seller_sku", "seller sku", "sku"],
    },
    "TikTok": {
        "order": ["Order ID", "order_id", "order id", "order no"],
        "sku":   ["Seller SKU", "seller_sku", "seller sku", "sku"],
    },
}

WMS_PRESET = {
    "client_order": ["client_order_id", "client order id", "clientorderid"],
    "sku":          ["sku", "sku_code", "seller_sku", "item_code"],
}


# ---------------------------------------------------------------- helpers
def read_file(uploaded, key):
    """Read CSV / Excel. For Excel with several tabs, let the user pick one."""
    if uploaded.name.lower().endswith(".csv"):
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


def pick(label, df, candidates, key, allow_none=False):
    """Selectbox that auto-selects the first matching column from candidates."""
    opts = ([NONE_OPT] if allow_none else []) + list(df.columns)
    lowered = [str(c).strip().lower() for c in opts]
    idx = 0
    for cand in candidates:
        if cand.strip().lower() in lowered:
            idx = lowered.index(cand.strip().lower())
            break
    return st.selectbox(label, opts, index=idx, key=key)


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
st.caption("Pick the marketplace, upload the sheets, confirm the columns and download the report.")

marketplace = st.selectbox("🛒 Marketplace", MARKETPLACES, index=0)

c1, c2, c3 = st.columns(3)
with c1:
    st.subheader("1️⃣ TC Orders sheet")
    tc_file = st.file_uploader("Required", type=["xlsx", "xls", "csv"], key="tc")
with c2:
    st.subheader(f"2️⃣ {marketplace} sheet")
    mp_file = st.file_uploader("Required", type=["xlsx", "xls", "csv"], key="mp")
with c3:
    st.subheader("3️⃣ WMS sheet")
    wms_file = st.file_uploader("Optional", type=["xlsx", "xls", "csv"], key="wms")

if not (tc_file and mp_file):
    st.info(f"Upload at least the TC Orders sheet and the {marketplace} sheet to begin.")
    st.stop()

# ---------------------------------------------------------------- column mapping
st.divider()
st.subheader("🔧 Column mapping")
st.caption("Columns are mapped automatically for the selected marketplace – change them only if needed.")
m1, m2, m3 = st.columns(3)
k = marketplace  # widget keys include marketplace so mapping refreshes when it changes

with m1:
    st.markdown("**TC Orders**")
    tc = read_file(tc_file, "tc")
    tc_order = pick("Order No", tc, TC_PRESET["order"], f"tc_order_{k}")
    tc_invoice = pick("Invoice No (used for WMS)", tc, TC_PRESET["invoice"], f"tc_inv_{k}", allow_none=True)
    tc_sku = pick("Seller SKU", tc, TC_PRESET["sku"], f"tc_sku_{k}", allow_none=True)
    tc_channel = pick("Channel (to filter marketplace)", tc, TC_PRESET["channel"], f"tc_ch_{k}", allow_none=True)

    # Keep only TC rows belonging to the selected marketplace
    tc_channels_sel = None
    if tc_channel != NONE_OPT:
        all_ch = sorted(v for v in tc[tc_channel].astype(str).str.strip().unique() if v)
        default_ch = [c for c in all_ch if marketplace.lower() in c.lower()]
        tc_channels_sel = st.multiselect(
            f"Channel values for {marketplace}", all_ch, default=default_ch, key=f"tc_chv_{k}"
        )

with m2:
    st.markdown(f"**{marketplace}**")
    mp = read_file(mp_file, "mp")
    mp_order = pick("Order No", mp, MP_PRESET[marketplace]["order"], f"mp_order_{k}")
    mp_sku = pick("Seller SKU", mp, MP_PRESET[marketplace]["sku"], f"mp_sku_{k}", allow_none=True)
    st.text_input("Marketplace Name", value=marketplace, disabled=True, key=f"mp_name_{k}")

wms = None
with m3:
    st.markdown("**WMS**")
    if wms_file:
        wms = read_file(wms_file, "wms")
        wms_client = pick("client_order_id", wms, WMS_PRESET["client_order"], f"wms_client_{k}")
        wms_sku = pick("SKU (optional)", wms, WMS_PRESET["sku"], f"wms_sku_{k}", allow_none=True)
        if tc_invoice == NONE_OPT:
            st.warning("Select the Invoice No column in TC Orders to compare with WMS.")
    else:
        st.caption("No WMS sheet uploaded – WMS check will be skipped.")

use_wms = wms is not None and tc_invoice != NONE_OPT

# ---------------------------------------------------------------- comparison
if st.button("▶️ Run order check", type="primary", use_container_width=True):
    tc_src = tc
    if tc_channels_sel is not None:
        if not tc_channels_sel:
            st.error(f"Select at least one channel value for {marketplace} in the TC Orders mapping.")
            st.stop()
        tc_src = tc[tc[tc_channel].astype(str).str.strip().isin(tc_channels_sel)]

    tcd = pd.DataFrame({
        "key": normalize(tc_src[tc_order]),
        "order": tc_src[tc_order].astype(str).str.strip(),
        "invoice": normalize(tc_src[tc_invoice]) if tc_invoice != NONE_OPT else "",
        "sku": tc_src[tc_sku].astype(str).str.strip() if tc_sku != NONE_OPT else "",
    })
    tcd = tcd[tcd["key"].ne("") & tcd["key"].ne("NAN")]

    mpd = pd.DataFrame({
        "key": normalize(mp[mp_order]),
        "order": mp[mp_order].astype(str).str.strip(),
        "sku": mp[mp_sku].astype(str).str.strip() if mp_sku != NONE_OPT else "",
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

    # one row per order, multiple SKUs joined
    tc_g = tcd.groupby("key").agg(order=("order", "first"), invoice=("invoice", join_unique),
                                  sku=("sku", join_unique)).reset_index()
    mp_g = mpd.groupby("key").agg(order=("order", "first"), sku=("sku", join_unique)).reset_index()
    merged = tc_g.merge(mp_g, on="key", how="outer", suffixes=("_tc", "_mp"))

    def val(x):
        return x if isinstance(x, str) and x else ""

    rows = []
    for _, r in merged.iterrows():
        in_tc, in_mp = r["key"] in tc_keys, r["key"] in mp_keys
        comments = []
        if not in_tc:
            comments.append("Order not available in TC Orders sheet")
        if not in_mp:
            comments.append(f"Order not available in {marketplace} sheet")
        if use_wms and in_tc:
            invoices = [i.strip() for i in str(r["invoice"]).split(",") if i.strip()]
            if not invoices:
                comments.append("Invoice No missing in TC – cannot check WMS")
            elif not any(i in wms_keys for i in invoices):
                comments.append(f"Invoice No {', '.join(invoices)} not available in WMS sheet")

        rows.append({
            "Marketplace Name": marketplace,
            "Order No": val(r["order_tc"]) or val(r["order_mp"]),
            "Seller SKU": val(r["sku_tc"]) or val(r["sku_mp"]),
            "Comments": "; ".join(comments) if comments else "Order available in all sheets",
        })

    # WMS records whose client_order_id doesn't match any TC invoice
    if use_wms:
        tc_invoices = set()
        for inv in tc_g["invoice"]:
            tc_invoices.update(i.strip() for i in str(inv).split(",") if i.strip())
        extra = wmsd[~wmsd["key"].isin(tc_invoices)].groupby("key").agg(
            raw=("raw", "first"), sku=("sku", join_unique)).reset_index()
        for _, r in extra.iterrows():
            rows.append({
                "Marketplace Name": marketplace,
                "Order No": r["raw"],
                "Seller SKU": r["sku"],
                "Comments": f"client_order_id {r['raw']} available in WMS but not in TC Orders (Invoice No)",
            })

    report = pd.DataFrame(rows, columns=["Marketplace Name", "Order No", "Seller SKU", "Comments"])
    ok_mask = report["Comments"] == "Order available in all sheets"
    summary = pd.DataFrame({
        "Metric": ["Marketplace", f"Orders in TC Orders sheet ({marketplace})",
                   f"Orders in {marketplace} sheet",
                   "Records in WMS sheet" if use_wms else "WMS check",
                   "Orders matched in all sheets", "Orders with issues"],
        "Value": [marketplace, len(tc_keys), len(mp_keys),
                  len(wms_keys) if use_wms else "Skipped",
                  int(ok_mask.sum()), int((~ok_mask).sum())],
    })
    st.session_state["result"] = (marketplace, report, summary)

# ---------------------------------------------------------------- results
if "result" in st.session_state and st.session_state["result"][0] == marketplace:
    _, report, summary = st.session_state["result"]
    ok_mask = report["Comments"] == "Order available in all sheets"

    st.divider()
    st.subheader(f"📊 Results – {marketplace}")
    kc = st.columns(4)
    kc[0].metric("TC orders", summary.iloc[1, 1])
    kc[1].metric(f"{marketplace} orders", summary.iloc[2, 1])
    kc[2].metric("✅ Matched", int(ok_mask.sum()))
    kc[3].metric("⚠️ Issues", int((~ok_mask).sum()))

    only_issues = st.toggle("Show only mismatches", value=True)
    search = st.text_input("🔍 Search order no / SKU")
    view = report[~ok_mask] if only_issues else report
    if search:
        s = search.strip().upper()
        view = view[view["Order No"].str.upper().str.contains(s, na=False, regex=False)
                    | view["Seller SKU"].str.upper().str.contains(s, na=False, regex=False)]
    st.dataframe(view, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download output sheet (Excel)",
        data=to_excel(report, summary),
        file_name=f"order_check_{marketplace.lower()}_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
