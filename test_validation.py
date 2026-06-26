import streamlit as st
import aiohttp
import asyncio
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

API_URL = "http://10.10.63.155:93/api/process/combinedOcrBytes"

CUSTOMERS = ["corybrothers_oil"]

COLUMNS = [
    ("Billing Document\nVBRK_VBELN",                                    "invoice_number",       "inv"),
    ("Due Date",                                                         "due_date",             "inv"),
    ("Document number of the\nreference document\nVBRP_VGBEL",          "delivery_note_number", "item"),
    ("Material Number",                                                  "material_number",      "item"),
    ("Material Name\nMAKT_MAKTX",                                       "description",          "item"),
    ("Billing quantity in\nstockkeeping unit\nVBRP_FKLMG",              "quantity",             "item"),
    ("UoM",                                                              "quantity_uom",         "item"),
    ("Net weight\nVBRP_NTGEW",                                          "nett_weight",          "item"),
    ("Gross weight\nVBRP_BRGEW",                                        "gross_weight",         "item"),
    ("Volume\nVBRP_VOLUM",                                              "volume",               "item"),
    ("Volume unit\nVBRP_VOLEH",                                         "volume_uom",           "item"),
    ("Net value of the billing item\nin document currency\nVBRP_NETWR", "line_net_price",       "item"),
    ("Rebate basis 1\nVBRP_BONBA",                                      "rebate_basis_amount",  "item"),
    ("SD Document Currency\nVBRK_WAERK",                                "currency",             "inv"),
    ("Country of Origin",                                               "country_of_origin",    "item"),
    ("Customs Tariff Number",                                           "hs_code",              "item"),
    ("Pref. Origin",                                                    "preferential_origin",  "item"),
    ("IncoTerms",                                                       "incoterms",            "item"),
    ("Unit Price",                                                      "unit_price",           "item"),
    ("Line Amount",                                                     "line_amount",          "item"),
]

# Short display labels for the UI table (no SAP codes)
UI_LABELS = [
    "Billing Doc", "Due Date", "Del. Note No.", "Material No.", "Material Name",
    "Qty", "UoM", "Net Weight", "Gross Weight", "Volume", "Vol. Unit",
    "Net Value", "Rebate Basis", "Currency", "Country of Origin",
    "Customs Tariff No.", "Pref. Origin", "IncoTerms", "Unit Price", "Line Amount",
]

COL_WIDTHS = [20, 14, 24, 16, 40, 10, 8, 12, 12, 10, 10, 18, 14, 12, 14, 18, 12, 18, 12, 14]


async def call_api(files_data: list, customer_name: str):
    try:
        data = aiohttp.FormData()
        for file_bytes, file_name in files_data:
            data.add_field("files", file_bytes, filename=file_name, content_type="application/pdf")
        data.add_field("customer_name", customer_name)
        data.add_field("model_name", "gpt-4o")
        async with aiohttp.ClientSession() as session:
            async with session.post(API_URL, data=data, timeout=aiohttp.ClientTimeout(total=300)) as response:
                return await response.json()
    except Exception as e:
        return {"error": str(e)}


def merge_delivery_into_invoices(invoices: list) -> list:
    lookup: dict = {}
    for inv in invoices:
        for item in inv.get("line_items", []):
            dn = item.get("delivery_note_number")
            seq = item.get("item_seq")
            if dn and seq:
                key = (str(dn), str(seq))
                if key not in lookup:
                    lookup[key] = item

    DR_FILL_FIELDS = [
        "material_number", "volume", "volume_uom",
        "nett_weight", "gross_weight", "hs_code",
        "country_of_origin", "preferential_origin", "incoterms", "quantity_uom",
    ]
    for inv in invoices:
        for item in inv.get("line_items", []):
            dn = item.get("delivery_note_number")
            seq = item.get("item_seq")
            if not dn or not seq:
                continue
            source = lookup.get((str(dn), str(seq)), {})
            for field in DR_FILL_FIELDS:
                if not item.get(field) and source.get(field):
                    item[field] = source[field]
    return invoices


def invoices_to_dataframe(invoices: list) -> pd.DataFrame:
    rows = []
    for inv in invoices:
        for item in inv.get("line_items", []):
            row = {}
            for label, field, source in zip(UI_LABELS, [c[1] for c in COLUMNS], [c[2] for c in COLUMNS]):
                val = inv.get(field, "") if source == "inv" else item.get(field, "")
                row[label] = "" if val is None else val
            rows.append(row)
    return pd.DataFrame(rows, columns=UI_LABELS)


def build_excel(invoices: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Data"

    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill("solid", fgColor="EEF2F7")
    inv_fill = PatternFill("solid", fgColor="D6E4F0")

    for col_idx, (header, _, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[1].height = 52

    row = 2
    for inv in invoices:
        items = inv.get("line_items", [])
        for i, item in enumerate(items):
            row_fill = inv_fill if i == 0 else (alt_fill if row % 2 == 0 else None)
            for col_idx, (_, field, source) in enumerate(COLUMNS, start=1):
                val = inv.get(field, "") if source == "inv" else item.get(field, "")
                if val is None:
                    val = ""
                cell = ws.cell(row=row, column=col_idx, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.alignment = left
                cell.border = border
                if row_fill:
                    cell.fill = row_fill
            row += 1

    for i, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ── PAGE CONFIG ───────────────────────────────────────────────────────────────

st.set_page_config(page_title="Invoice Extractor", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .stApp { background: #F7F9FC; }
    .main-title { font-size: 26px; font-weight: 700; color: #1F3864; margin-bottom: 2px; }
    .sub-title { font-size: 14px; color: #6B7280; margin-bottom: 20px; }

    .stat-card {
        background: white; border-radius: 10px; padding: 14px 18px;
        border: 1px solid #E5E7EB; text-align: center;
    }
    .stat-num { font-size: 26px; font-weight: 700; color: #1F3864; }
    .stat-label { font-size: 11px; color: #6B7280; margin-top: 2px; }

    .file-pill {
        display:inline-block; background:#EEF2F7; color:#1F3864;
        border-radius:20px; padding:3px 12px; font-size:12px;
        margin:3px 4px 3px 0; font-weight:500;
    }
    .file-pill.dr { background:#FEF3C7; color:#92400E; }

    div[data-testid="stFileUploader"] {
        background: white; border-radius: 10px;
        border: 1.5px dashed #CBD5E1; padding: 12px;
    }

    /* Excel-style table viewer */
    .excel-wrap {
        border-radius: 10px; overflow: hidden;
        border: 1px solid #D1D5DB;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .excel-scroll {
        overflow-x: auto; overflow-y: auto;
        max-height: 520px;
    }
    .excel-table {
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        font-size: 12px;
        white-space: nowrap;
        min-width: 100%;
    }
    .excel-table thead th {
        background: #1F3864;
        color: white;
        padding: 8px 12px;
        text-align: center;
        font-weight: 600;
        font-size: 11px;
        position: sticky;
        top: 0;
        z-index: 2;
        border-right: 1px solid #2d4d7a;
        border-bottom: 2px solid #152849;
    }
    .excel-table thead th:first-child { position: sticky; left: 0; z-index: 3; }
    .excel-table tbody td {
        padding: 6px 12px;
        border-right: 1px solid #E5E7EB;
        border-bottom: 1px solid #E5E7EB;
        color: #111827;
    }
    .excel-table tbody td:first-child {
        position: sticky; left: 0; z-index: 1;
        font-weight: 600; color: #1F3864;
        background: inherit;
    }
    .excel-table tbody tr:nth-child(odd)  { background: #FFFFFF; }
    .excel-table tbody tr:nth-child(even) { background: #EEF2F7; }
    .excel-table tbody tr.inv-first { background: #D6E4F0 !important; }
    .excel-table tbody tr:hover td { background: #DBEAFE !important; }

    .badge-yes { display:inline-block; background:#D1FAE5; color:#065F46;
        border-radius:10px; padding:1px 8px; font-size:11px; font-weight:600; }
    .badge-no  { display:inline-block; background:#FEE2E2; color:#991B1B;
        border-radius:10px; padding:1px 8px; font-size:11px; font-weight:600; }

    .section-title {
        font-size: 15px; font-weight: 600; color: #1F3864;
        margin: 20px 0 10px 0; display: flex; align-items: center; gap: 8px;
    }
    .row-count {
        font-size: 12px; font-weight: 400; color: #6B7280;
        background: #F3F4F6; border-radius: 20px; padding: 1px 10px;
    }
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────────────────────

st.markdown('<div class="main-title">📄 Invoice Data Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Upload invoices and delivery receipts — data is merged and shown in SAP format.</div>', unsafe_allow_html=True)

col_up, col_cust = st.columns([3, 2])
with col_up:
    uploaded_files = st.file_uploader(
        "Upload PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
with col_cust:
    customer = st.selectbox("Customer", CUSTOMERS, index=0)

if uploaded_files:
    pills_html = ""
    for f in uploaded_files:
        is_dr = any(k in f.name.lower() for k in ["delivery", "receipt", "dr", "dn", "despatch"])
        cls = "file-pill dr" if is_dr else "file-pill"
        icon = "📦" if is_dr else "🧾"
        pills_html += f'<span class="{cls}">{icon} {f.name}</span>'
    st.markdown(pills_html, unsafe_allow_html=True)

    st.divider()
    if st.button("⚡ Extract & Merge Data", use_container_width=True, type="primary"):
        with st.spinner(f"Processing {len(uploaded_files)} file(s) through OCR pipeline..."):
            files_data = [(f.getvalue(), f.name) for f in uploaded_files]
            result = asyncio.run(call_api(files_data, customer))

        if "error" in result:
            st.error(f"API error: {result['error']}")
            st.stop()

        try:
            invoices = result["data"]["extracted_data"]["gpt_extraction_output"]["invoices"]
        except (KeyError, TypeError):
            st.error("Unexpected response format from API.")
            with st.expander("Raw API response"):
                st.json(result)
            st.stop()

        invoices = merge_delivery_into_invoices(invoices)
        st.session_state["invoices"] = invoices
        st.session_state["uploaded_names"] = [f.name for f in uploaded_files]
        st.session_state["n_files"] = len(uploaded_files)

# ── RESULTS ───────────────────────────────────────────────────────────────────

if "invoices" in st.session_state:
    invoices = st.session_state["invoices"]
    total_items = sum(len(inv.get("line_items", [])) for inv in invoices)
    total_value = sum(inv.get("net_price_total", 0) or 0 for inv in invoices)
    n_files = st.session_state.get("n_files", len(uploaded_files) if uploaded_files else 0)

    # ── Stats ─────────────────────────────────────────────────────────────────
    st.markdown("### Summary")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-card"><div class="stat-num">{n_files}</div><div class="stat-label">Files uploaded</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-card"><div class="stat-num">{len(invoices)}</div><div class="stat-label">Invoices found</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="stat-card"><div class="stat-num">{total_items}</div><div class="stat-label">Line items</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="stat-card"><div class="stat-num">€{total_value:,.2f}</div><div class="stat-label">Total net value</div></div>', unsafe_allow_html=True)

    # ── Excel-style table ─────────────────────────────────────────────────────
    st.markdown(
        f'<div class="section-title">📊 Data preview <span class="row-count">{total_items} rows · {len(COLUMNS)} columns</span></div>',
        unsafe_allow_html=True,
    )

    # Build HTML table
    th_cells = "".join(f"<th>{lbl.replace(chr(10), '<br>')}</th>" for lbl in UI_LABELS)
    tbody_rows = ""

    for inv in invoices:
        items = inv.get("line_items", [])
        for i, item in enumerate(items):
            row_class = "inv-first" if i == 0 else ""
            tds = ""
            for label, (_, field, source) in zip(UI_LABELS, COLUMNS):
                val = inv.get(field, "") if source == "inv" else item.get(field, "")
                if val is None:
                    val = ""

                # Special rendering
                if label == "Pref. Origin":
                    if str(val).upper() == "YES":
                        cell_html = '<span class="badge-yes">YES</span>'
                    elif str(val).upper() == "NO":
                        cell_html = '<span class="badge-no">NO</span>'
                    else:
                        cell_html = str(val)
                elif label in ("Net Value", "Unit Price", "Line Amount", "Rebate Basis"):
                    cell_html = f"{float(val):,.2f}" if val != "" else ""
                elif label in ("Net Weight", "Gross Weight", "Volume", "Qty"):
                    cell_html = f"{float(val):,.3f}" if val != "" else ""
                else:
                    cell_html = str(val)

                tds += f"<td>{cell_html}</td>"
            tbody_rows += f"<tr class='{row_class}'>{tds}</tr>"

    table_html = f"""
    <div class="excel-wrap">
      <div class="excel-scroll">
        <table class="excel-table">
          <thead><tr>{th_cells}</tr></thead>
          <tbody>{tbody_rows}</tbody>
        </table>
      </div>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)

    # ── Filter + Download row ──────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    col_filter, col_dl = st.columns([3, 2])

    with col_filter:
        inv_nums = ["All invoices"] + [inv.get("invoice_number", "?") for inv in invoices]
        selected = st.selectbox("Filter by invoice", inv_nums, label_visibility="collapsed")

    with col_dl:
        # Build filtered or full excel
        if selected == "All invoices":
            export_invoices = invoices
            fname_base = st.session_state.get("uploaded_names", ["invoices"])[0].replace(".pdf", "")
        else:
            export_invoices = [inv for inv in invoices if inv.get("invoice_number") == selected]
            fname_base = f"invoice_{selected}"

        excel_bytes = build_excel(export_invoices)
        st.download_button(
            label="⬇️ Download Excel (SAP format)",
            data=excel_bytes,
            file_name=f"{fname_base}_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    # Apply filter — re-render table if filtered
    if selected != "All invoices":
        filtered = [inv for inv in invoices if inv.get("invoice_number") == selected]
        filtered_items = sum(len(inv.get("line_items", [])) for inv in filtered)
        st.caption(f"Showing {filtered_items} line item(s) for invoice {selected}")

        tbody_filtered = ""
        for inv in filtered:
            for i, item in enumerate(inv.get("line_items", [])):
                row_class = "inv-first" if i == 0 else ""
                tds = ""
                for label, (_, field, source) in zip(UI_LABELS, COLUMNS):
                    val = inv.get(field, "") if source == "inv" else item.get(field, "")
                    if val is None:
                        val = ""
                    if label == "Pref. Origin":
                        if str(val).upper() == "YES":
                            cell_html = '<span class="badge-yes">YES</span>'
                        elif str(val).upper() == "NO":
                            cell_html = '<span class="badge-no">NO</span>'
                        else:
                            cell_html = str(val)
                    elif label in ("Net Value", "Unit Price", "Line Amount", "Rebate Basis"):
                        cell_html = f"{float(val):,.2f}" if val != "" else ""
                    elif label in ("Net Weight", "Gross Weight", "Volume", "Qty"):
                        cell_html = f"{float(val):,.3f}" if val != "" else ""
                    else:
                        cell_html = str(val)
                    tds += f"<td>{cell_html}</td>"
                tbody_filtered += f"<tr class='{row_class}'>{tds}</tr>"

        st.markdown(f"""
        <div class="excel-wrap">
          <div class="excel-scroll">
            <table class="excel-table">
              <thead><tr>{th_cells}</tr></thead>
              <tbody>{tbody_filtered}</tbody>
            </table>
          </div>
        </div>
        """, unsafe_allow_html=True)
