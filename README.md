# 📦 Order Checker (Streamlit)

Compares orders between the **TC Orders** sheet, the **Marketplace** sheet and (optionally) the **WMS** sheet, and produces an output sheet with:

| Marketplace Name | Order No | Seller SKU | Comments |
|---|---|---|---|

## Matching logic
- **TC Orders ↔ Marketplace**: matched on **Order No** (VLOOKUP-style, both directions).
- **TC Orders ↔ WMS** (optional): TC **Invoice No** is matched against WMS **client_order_id**.
- Keys are cleaned before matching (spaces trimmed, case ignored, trailing `.0` removed).
- Comments examples:
  - `Order available in all sheets`
  - `Order not available in Marketplace sheet`
  - `Order not available in TC Orders sheet`
  - `Invoice No XXXX not available in WMS sheet`
  - `client_order_id XXXX available in WMS but not in TC Orders (Invoice No)`

The downloaded Excel has 3 tabs: **Summary**, **Order Check** (all orders), **Mismatches Only**.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (free) on Streamlit Community Cloud
1. Push this folder to a GitHub repo.
2. Go to https://share.streamlit.io → **New app** → pick the repo, branch `main`, file `app.py`.
3. Click **Deploy** and share the link with your team.
