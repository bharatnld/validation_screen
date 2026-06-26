import streamlit as st
import pandas as pd
import asyncio
import json
import io
import re
from streamlit_pdf_viewer import pdf_viewer          # ← added
from api_client import process_file               # ← keep your original import
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom


st.set_page_config(
    page_title="OCR Validation Portal",
    layout="wide"
)

# ===========================
# CSS
# ===========================

st.markdown("""
<style>

.block-container{
    padding-top:0.5rem;
    max-width:100%;
}

.section-box{
    border:1px solid #dcdcdc;
    padding:10px;
    border-radius:8px;
    margin-bottom:10px;
}

.order-header{
    font-size:28px;
    font-weight:bold;
    color:#2d89ef;
}

</style>
""", unsafe_allow_html=True)

# ===========================
# HELPERS
# ===========================
def normalize(value):
    """Text normalization: case/whitespace-insensitive, and strips trailing
    punctuation like a stray comma or period (e.g. 'Rakem LTD,' vs 'Rakem LTD')."""
    if value is None:
        return ""
    s = str(value).strip()
    s = s.strip(" ,.;:")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def parse_number(value):
    """
    Try to read a string as a number, tolerating both
    '1,234.56' (comma=thousands, point=decimal) and
    '1.234,56' / '24141,6' (point=thousands, comma=decimal) styles,
    plus pure thousands-grouping with no decimal part at all
    (e.g. '13,906,000' or '13.906.000'). Returns a float, or None
    if the string isn't number-like.
    """
    if value is None:
        return None

    s = str(value).strip()
    if s == "":
        return None

    # drop anything that isn't a digit, separator, or sign (units like "kg" etc.)
    s = re.sub(r"[^0-9.,\-]", "", s)
    if s in ("", "-", ".", ","):
        return None

    has_comma = "," in s
    has_dot = "." in s

    if has_comma and has_dot:
        # whichever separator appears LAST is the real decimal point
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        last_group = s.split(",")[-1]
        if s.count(",") == 1 and len(last_group) in (1, 2):
            # one comma, 1-2 trailing digits -> decimal separator: "24141,6"
            s = s.replace(",", ".")
        else:
            # multiple commas, or 3 trailing digits -> thousands grouping: "13,906,000"
            s = s.replace(",", "")
    elif has_dot:
        last_group = s.split(".")[-1]
        if s.count(".") > 1 or len(last_group) == 3:
            # multiple dots, or exactly 3 trailing digits -> thousands grouping
            s = s.replace(".", "")
        # else: single dot with 1-2 trailing digits is already a normal decimal point

    try:
        return float(s)
    except ValueError:
        return None


def values_match(ai_value, gt_value, tolerance=0.01):
    """
    Compares two field values the way a human would: text is matched after
    normalizing case/whitespace/trailing punctuation, and number-like values
    are matched numerically so formatting differences (thousands separators,
    decimal comma vs point, trailing zeros: '13,906.000' vs '13906') don't
    register as a mismatch.
    """
    ai_str = safe_str(ai_value)
    gt_str = safe_str(gt_value)

    if normalize(ai_str) == normalize(gt_str):
        return True

    ai_num = parse_number(ai_str)
    gt_num = parse_number(gt_str)

    if ai_num is not None and gt_num is not None:
        return abs(ai_num - gt_num) <= tolerance

    return False


def safe_get(d, key, default=""):
    """
    Dict.get() that also treats an *explicit* null/None value the same
    as a missing key. Plain `.get(key, default)` only falls back when
    the key is absent - if the API returns `"field": null`, it still
    returns None, which then leaks into the UI/XML as the literal
    string "None". This also avoids the `value or ""` bug, which
    incorrectly blanks out legitimate falsy values like 0.
    """
    if d is None:
        return default
    val = d.get(key, default)
    return default if val is None else val


def safe_dict(d, key):
    """Get a nested dict value, defaulting to {} if missing, None, or not a dict."""
    if d is None:
        return {}
    val = d.get(key)
    return val if isinstance(val, dict) else {}


def safe_str(value):
    """Coerce a value to a string for XML text content (ET requires str, not int/float/None)."""
    return "" if value is None else str(value)


def parse_ground_truth_xml(xml_path):

    tree = ET.parse(xml_path)
    root = tree.getroot()

    mrn = root.find(".//MRN")

    if mrn is None:
        return {}

    consignor = mrn.find("Consignor")
    consignee = mrn.find("Consignee")

    good = mrn.find(".//Good")

    return {

        "MRN": mrn.findtext("Number", ""),
        "Type": mrn.findtext("Type", ""),

        "Exporter Name":
            consignor.findtext("Name", "") if consignor is not None else "",

        "Exporter Address":
            consignor.findtext("Address", "") if consignor is not None else "",

        "Exporter City":
            consignor.findtext("City", "") if consignor is not None else "",

        "Exporter Postal":
            consignor.findtext("Postalcode", "") if consignor is not None else "",

        "Exporter Country":
            consignor.findtext("Country", "") if consignor is not None else "",

        "Exporter EORI":
            consignor.findtext("EoriNumber", "") if consignor is not None else "",

        "Importer Name":
            consignee.findtext("Name", "") if consignee is not None else "",

        "Importer Address":
            consignee.findtext("Address", "") if consignee is not None else "",

        "Importer City":
            consignee.findtext("City", "") if consignee is not None else "",

        "Importer Postal":
            consignee.findtext("Postalcode", "") if consignee is not None else "",

        "Importer Country":
            consignee.findtext("Country", "") if consignee is not None else "",

        "Importer EORI":
            consignee.findtext("EoriNumber", "") if consignee is not None else "",

        "TotalPackages":
            mrn.findtext("TotalPackages", ""),

        "GrossMass":
            mrn.findtext("GrossMass", ""),

        # Searched from root with ".//" rather than as a direct child of
        # <MRN>, since these tags can appear at different nesting depths
        # depending on the source XML - this way it's found either way.
        "Destination Office Text":
            root.findtext(".//DestinationOfficeText", ""),

        "Destination Office Country":
            root.findtext(".//DestinationOfficeCountry", ""),

        "CommodityCode":
            good.findtext("CommodityCode", "") if good is not None else "",

        "GoodDescription":
            good.findtext("GoodDescription", "") if good is not None else ""
    }


def extract_ai_fields(result):

    shipment = safe_dict(result, "shipment_metadata")
    exporter = safe_dict(result, "exporter_information")
    importer = safe_dict(result, "importer_information")
    items = result.get("items") or []

    item = items[0] if items else {}

    exp_addr = safe_dict(exporter, "address")
    imp_addr = safe_dict(importer, "address")

    return {

        "MRN": safe_get(shipment, "MRN"),
        "Type": safe_get(shipment, "customtype"),

        "Exporter Name": safe_get(exporter, "name"),
        "Exporter Address": safe_get(exp_addr, "street_and_nr"),
        "Exporter City": safe_get(exp_addr, "city"),
        "Exporter Postal": safe_get(exp_addr, "postal_code"),
        "Exporter Country": safe_get(exp_addr, "country"),
        "Exporter EORI": safe_get(exporter, "tax_id"),

        "Importer Name": safe_get(importer, "name"),
        "Importer Address": safe_get(imp_addr, "street_and_nr"),
        "Importer City": safe_get(imp_addr, "city"),
        "Importer Postal": safe_get(imp_addr, "postal_code"),
        "Importer Country": safe_get(imp_addr, "country"),
        "Importer EORI": safe_get(importer, "tax_id"),

        "TotalPackages":
            safe_get(shipment, "total_quantity"),

        "GrossMass":
            safe_get(shipment, "total_gross_weight"),

        "Destination Office Text":
            safe_get(shipment, "destination_office_text"),

        "Destination Office Country":
            safe_get(shipment, "destination_office_country"),

        "CommodityCode":
            safe_get(item, "hs_code"),

        "GoodDescription":
            safe_get(item, "description")
    }


def json_to_xml(result):

    shipment = safe_dict(result, "shipment_metadata")
    exporter = safe_dict(result, "exporter_information")
    importer = safe_dict(result, "importer_information")
    items = result.get("items") or []

    item = items[0] if items else {}

    root = ET.Element("MRNResponse")

    body = ET.SubElement(root, "Body")

    document = ET.SubElement(body, "Document")

    mrns = ET.SubElement(document, "MRNs")

    mrn = ET.SubElement(mrns, "MRN")

    ET.SubElement(mrn, "Number").text = safe_str(safe_get(shipment, "MRN"))
    ET.SubElement(mrn, "Type").text = safe_str(safe_get(shipment, "customtype"))

    consignor = ET.SubElement(mrn, "Consignor")

    exp_addr = safe_dict(exporter, "address")

    ET.SubElement(consignor, "Name").text = safe_str(safe_get(exporter, "name"))
    ET.SubElement(consignor, "Address").text = safe_str(safe_get(exp_addr, "street_and_nr"))
    ET.SubElement(consignor, "City").text = safe_str(safe_get(exp_addr, "city"))
    ET.SubElement(consignor, "Postalcode").text = safe_str(safe_get(exp_addr, "postal_code"))
    ET.SubElement(consignor, "Country").text = safe_str(safe_get(exp_addr, "country"))
    ET.SubElement(consignor, "EoriNumber").text = safe_str(safe_get(exporter, "tax_id"))

    consignee = ET.SubElement(mrn, "Consignee")

    imp_addr = safe_dict(importer, "address")

    ET.SubElement(consignee, "Name").text = safe_str(safe_get(importer, "name"))
    ET.SubElement(consignee, "Address").text = safe_str(safe_get(imp_addr, "street_and_nr"))
    ET.SubElement(consignee, "City").text = safe_str(safe_get(imp_addr, "city"))
    ET.SubElement(consignee, "Postalcode").text = safe_str(safe_get(imp_addr, "postal_code"))
    ET.SubElement(consignee, "Country").text = safe_str(safe_get(imp_addr, "country"))
    ET.SubElement(consignee, "EoriNumber").text = safe_str(safe_get(importer, "tax_id"))

    ET.SubElement(mrn, "TotalPackages").text = safe_str(safe_get(shipment, "total_quantity"))
    ET.SubElement(mrn, "GrossMass").text = safe_str(safe_get(shipment, "total_gross_weight"))
    ET.SubElement(mrn, "DestinationOfficeText").text = safe_str(safe_get(shipment, "destination_office_text"))
    ET.SubElement(mrn, "DestinationOfficeCountry").text = safe_str(safe_get(shipment, "destination_office_country"))

    goods = ET.SubElement(mrn, "Goods")

    good = ET.SubElement(goods, "Good")

    ET.SubElement(good, "CommodityCode").text = safe_str(safe_get(item, "hs_code"))
    ET.SubElement(good, "GoodDescription").text = safe_str(safe_get(item, "description"))

    xml_str = ET.tostring(root, encoding="utf-8")

    return minidom.parseString(xml_str).toprettyxml(indent="  ")


def compare_fields(ai_data, gt_data):

    rows = []

    matched = 0

    total = len(gt_data)

    for field in gt_data:

        ai_value = safe_str(ai_data.get(field, ""))
        gt_value = safe_str(gt_data.get(field, ""))

        is_match = values_match(ai_value, gt_value)

        if is_match:
            matched += 1

        rows.append({
            "Field": field,
            "AI Value": ai_value,
            "Ground Truth": gt_value,
            "Match": "✅" if is_match else "❌"
        })

    accuracy = round((matched / total) * 100, 2) if total else 0.0

    return pd.DataFrame(rows), accuracy


def to_excel_bytes(sheets: dict):
    """
    Write one or more DataFrames to an in-memory .xlsx file.
    `sheets` is a dict of {sheet_name: DataFrame}. Returns raw bytes,
    suitable for st.download_button. Requires openpyxl
    (pip install openpyxl).
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return buffer.getvalue()


async def _process_one_batch_file(uploaded, gt_folder, customer_name, model_name, semaphore, progress_state, progress_callback):
    """Run OCR + comparison for a single uploaded PDF, guarded by `semaphore`
    so only a limited number of files hit the AI API at the same time."""

    pdf_name = uploaded.name

    async with semaphore:

        xml_name = os.path.splitext(pdf_name)[0] + ".xml"
        xml_path = os.path.join(gt_folder, xml_name)

        if not os.path.exists(xml_path):
            summary_row = {
                "File": pdf_name,"Model": model_name,  "Status": "Missing ground truth XML",
                "Accuracy %": None, "Matched Fields": None, "Total Fields": None
            }
            detail_rows = []

        else:
            try:
                uploaded.seek(0)  # rewind in case it was read before (e.g. preview)
                response = await process_file(uploaded, customer_name, model_name)

                result = get_extracted_data(response)
                ai_data = extract_ai_fields(result)
                gt_data = parse_ground_truth_xml(xml_path)

                if not gt_data:
                    summary_row = {
                        "File": pdf_name, "Model": model_name, "Status": "No <MRN> in ground truth XML",
                        "Accuracy %": None, "Matched Fields": None, "Total Fields": None
                    }
                    detail_rows = []
                else:
                    compare_df, accuracy = compare_fields(ai_data, gt_data)
                    matched = int((compare_df["Match"] == "✅").sum())
                    total = len(compare_df)

                    summary_row = {
                        "File": pdf_name, "Status": "OK","Model": model_name,
                        "Accuracy %": accuracy, "Matched Fields": matched, "Total Fields": total
                    }
                    detail_rows = [
                        {
                            "File": pdf_name,
                            "Field": row["Field"],
                            "AI Value": row["AI Value"],
                            "Ground Truth": row["Ground Truth"],
                            "Match": row["Match"]
                        }
                        for _, row in compare_df.iterrows()
                    ]

            except Exception as exc:
                summary_row = {
                    "File": pdf_name, "Model": model_name, "Status": f"OCR error: {exc}",
                    "Accuracy %": None, "Matched Fields": None, "Total Fields": None
                }
                detail_rows = []

        progress_state["done"] += 1
        if progress_callback:
            progress_callback(progress_state["done"], progress_state["total"], pdf_name)

        return summary_row, detail_rows


async def _run_batch_accuracy_async(uploaded_pdfs, gt_folder, customer_name, model_name, max_concurrency, progress_callback):

    semaphore = asyncio.Semaphore(max_concurrency)
    progress_state = {"done": 0, "total": len(uploaded_pdfs)}

    tasks = [
        _process_one_batch_file(uploaded, gt_folder, customer_name, model_name, semaphore, progress_state, progress_callback)
        for uploaded in uploaded_pdfs
    ]

    results = await asyncio.gather(*tasks)

    summary_rows = [r[0] for r in results]
    detail_rows = [row for r in results for row in r[1]]

    return pd.DataFrame(summary_rows), pd.DataFrame(detail_rows)


def run_batch_accuracy(uploaded_pdfs, gt_folder, customer_name, model_name, progress_callback=None, max_concurrency=4):
    """
    Process every uploaded PDF (from st.file_uploader, accept_multiple_files=True),
    match each one to a same-named .xml file in `gt_folder` (e.g. ABC.pdf ↔ ABC.xml),
    run OCR + comparison CONCURRENTLY (up to `max_concurrency` files hitting the
    AI API at once) and return (summary_df, detail_df, error_message).

    summary_df: one row per PDF -> File, Status, Accuracy %, Matched Fields, Total Fields
    detail_df:  one row per PDF per field -> File, Field, AI Value, Ground Truth, Match

    NOTE: this assumes process_file()'s underlying HTTP client supports
    concurrent calls (e.g. a per-call session, or an httpx/aiohttp client
    safe for concurrent use). If api_client.py shares one global session/
    connection that isn't concurrency-safe, lower max_concurrency to 1 or
    fix that client first.
    """

    if not uploaded_pdfs:
        return None, None, "No PDF files uploaded."

    if not os.path.isdir(gt_folder):
        return None, None, f"Ground truth folder not found: {gt_folder}"

    summary_df, detail_df = asyncio.run(
        _run_batch_accuracy_async(uploaded_pdfs, gt_folder, customer_name, model_name, max_concurrency, progress_callback)
    )
    return summary_df, detail_df, None


def get_extracted_data(response):
    return (
        (response or {})
        .get("data", {})
        .get("extracted_data", {})
        .get("gpt_extraction_output", {})
    ) or {}

# ===========================
# HEADER
# ===========================

st.markdown(
    '<div class="order-header">OCR Validation Portal</div>',
    unsafe_allow_html=True
)

st.divider()

# ===========================
# TOP BAR
# ===========================


col1, col2, col3, col4 = st.columns([3, 1, 1, 1])

with col1:
    uploaded_file = st.file_uploader(
        "Upload PDF",
        type=["pdf"]
    )

    # ── FIX: read bytes immediately into session_state so the cursor
    #         is not exhausted by the time the preview runs ──────────
    if uploaded_file is not None:
        st.session_state["pdf_bytes"] = uploaded_file.read()
        uploaded_file.seek(0)          # reset so process_file can read it too

with col2:
    customer_name = st.selectbox(
        "Customer",
        [
            "smeetferrybol",
            "smeetferryead",
            "irf_bol"
        ]
    )

with col3:
    process_btn = st.button("🚀 Process")


with col4:
    model_name = st.selectbox(
        "AI Model",
        [
            "gemini-3.1-pro-preview",
            "gemini-3.5-flash",
            "gemini-2.0-flash",
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-1.5-pro",
            "gemini-1.5-flash",
        ]
    )

# ===========================
# PROCESS FILE
# ===========================

if process_btn and uploaded_file:

    with st.spinner("Running OCR..."):


        response = asyncio.run(process_file(uploaded_file, customer_name, model_name))

        st.session_state["api_response"] = response
        st.session_state["model_used"] = model_name

# ===========================
# DISPLAY DATA
# ===========================

if "api_response" in st.session_state:

    response = st.session_state["api_response"]
    result   = get_extracted_data(response)

    shipment = safe_dict(result, "shipment_metadata")
    exporter = safe_dict(result, "exporter_information")
    importer = safe_dict(result, "importer_information")
    items    = result.get("items") or []

    tab1, tab2 = st.tabs(
        [
            "Validation",
            "Accuracy Report"
        ]
    )

    with tab1:
        left, right = st.columns([1, 2])
    # ==================================
    # LEFT PANEL
    # ==================================

        with left:

            with st.expander("SHIPMENT METADATA", expanded=True):

                shipment["customtype"] = st.text_input(
                    "Custom Type", safe_get(shipment, "customtype"))

                shipment["MRN"] = st.text_input(
                    "MRN", safe_get(shipment, "MRN"))

                shipment["total_quantity"] = st.text_input(
                    "Total Quantity", safe_str(safe_get(shipment, "total_quantity")))

                shipment["total_gross_weight"] = st.text_input(
                    "Gross Weight", safe_str(safe_get(shipment, "total_gross_weight")))

                shipment["date"] = st.text_input(
                    "Date", safe_get(shipment, "date"))

                shipment["destination_office_text"] = st.text_input(
                    "Destination Office Text", safe_get(shipment, "destination_office_text"))

                shipment["destination_office_country"] = st.text_input(
                    "Destination Office Country", safe_get(shipment, "destination_office_country"))

                shipment["shipment_reference"] = st.text_input(
                    "Shipment Reference", safe_get(shipment, "shipment_reference"))


            with st.expander("EXPORTER", expanded=True):

                exporter["name"] = st.text_input(
                    "Exporter Name", safe_get(exporter, "name"))

                exporter["tax_id"] = st.text_input(
                    "Exporter Tax ID", safe_get(exporter, "tax_id"))

                address = safe_dict(exporter, "address")

                address["street_and_nr"] = st.text_input(
                    "Street", safe_get(address, "street_and_nr"))
                address["city"]          = st.text_input(
                    "City",   safe_get(address, "city"))
                address["country"]       = st.text_input(
                    "Country", safe_get(address, "country"))
                address["postal_code"]   = st.text_input(
                    "Postal Code", safe_get(address, "postal_code"))

                exporter["address"] = address

            with st.expander("IMPORTER", expanded=True):

                importer["name"] = st.text_input(
                    "Importer Name", safe_get(importer, "name"))

                importer["tax_id"] = st.text_input(
                    "Importer Tax ID", safe_get(importer, "tax_id"))

                address = safe_dict(importer, "address")

                address["street_and_nr"] = st.text_input(
                    "Street ", safe_get(address, "street_and_nr"))
                address["city"]          = st.text_input(
                    "City ",   safe_get(address, "city"))
                address["country"]       = st.text_input(
                    "Country ", safe_get(address, "country"))
                address["postal_code"]   = st.text_input(
                    "Postal Code ", safe_get(address, "postal_code"))

                importer["address"] = address

        # ==================================
        # PDF PREVIEW
        # ==================================

        with right:

            st.subheader("PDF Preview")

            pdf_bytes = st.session_state.get("pdf_bytes")

            if pdf_bytes:

                st.download_button(
                    "📄 Download PDF",
                    pdf_bytes,
                    file_name=uploaded_file.name if uploaded_file else "document.pdf"
                )

                # ── Embedded viewer ──────────────────────────────────────
                pdf_viewer(
                    input=pdf_bytes,
                    width=700,
                    height=800,
                    render_text=True          # makes text selectable
                )
                # ────────────────────────────────────────────────────────

            else:
                st.info("Upload a PDF above to see the preview here.")

        # ==================================
        # ITEMS GRID
        # ==================================

        st.divider()
        st.subheader("LINE ITEMS")

        if items:

            df = pd.DataFrame(items)

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True
            )

            result["items"] = edited_df.to_dict(orient="records")

        # ==================================
        # SAVE RESULT
        # ==================================

        result["shipment_metadata"]    = shipment
        result["exporter_information"] = exporter
        result["importer_information"] = importer

        st.divider()

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button("✅ Approve"):
                with open("validated_output.json", "w", encoding="utf-8") as f:
                    json.dump(result, f, indent=4, ensure_ascii=False)
                st.success("Validation Approved")

        with c2:
            if st.button("❌ Reject"):
                st.error("Validation Rejected")

        with c3:
            st.download_button(
                "📥 Download JSON",
                json.dumps(result, indent=4, ensure_ascii=False),
                file_name="validated_output.json",
                mime="application/json"
            )
    with tab2:

        st.header("Accuracy Report")

        if uploaded_file:

            pdf_name = uploaded_file.name

            xml_name = os.path.splitext(pdf_name)[0] + ".xml"

            xml_path = os.path.join(
                "ground_truth",
                xml_name
            )

            st.write("PDF:", pdf_name)
            st.write("Expected XML:", xml_name)

            if os.path.exists(xml_path):

                gt_data = parse_ground_truth_xml(xml_path)

                if not gt_data:
                    st.error(
                        "Ground truth XML was found but no <MRN> element "
                        "could be parsed from it."
                    )
                else:

                    ai_data = extract_ai_fields(result)

                    compare_df, accuracy = compare_fields(
                        ai_data,
                        gt_data
                    )

                    st.metric(
                        "Overall Accuracy",
                        f"{accuracy}%"
                    )
                    col_acc, col_model = st.columns(2)
                    with col_acc:
                        st.metric("Overall Accuracy", f"{accuracy}%")
                    with col_model:
                        st.metric("Model Used", st.session_state.get("model_used", "N/A"))

                    st.dataframe(
                        compare_df,
                        use_container_width=True
                    )

                    generated_xml = json_to_xml(result)

                    st.download_button(
                        "📥 Download Generated XML",
                        generated_xml,
                        file_name=xml_name,
                        mime="application/xml"
                    )

                    accuracy_excel = to_excel_bytes({
                        "Summary": pd.DataFrame([{
                            "File": pdf_name,
                            "Model": st.session_state.get("model_used", "N/A"),  # ← added
                            "Accuracy %": accuracy,
                            "Matched Fields": int((compare_df["Match"] == "✅").sum()),
                            "Total Fields": len(compare_df)
                        }]),
                        "Comparison": compare_df
                    })

                    st.download_button(
                        "📊 Download Accuracy Report (Excel)",
                        accuracy_excel,
                        file_name=os.path.splitext(pdf_name)[0] + "_accuracy.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            else:

                st.error(
                    f"Ground truth XML not found: {xml_path}"
                )
        else:
            st.info("Upload and process a PDF to see the accuracy report.")

# ===========================
# BATCH ACCURACY REPORT
# (upload multiple PDFs at once, match each one to its ground-truth XML
# by filename - ABC.pdf ↔ ABC.xml - and export one combined Excel report)
# ===========================

st.divider()

with st.expander("📊 Batch Accuracy Report (multi-PDF upload)", expanded=False):

    st.caption(
        "Upload several PDFs at once. Each one is matched to a ground "
        "truth XML with the same base name in the folder below "
        "(e.g. ABC.pdf ↔ ABC.xml)."
    )

    batch_pdfs = st.file_uploader(
        "Upload PDFs for batch accuracy check",
        type=["pdf"],
        accept_multiple_files=True,
        key="batch_pdf_uploader"
    )

    bcol1, bcol2, bcol3, bcol4 = st.columns(4)

    with bcol1:
        gt_folder_input = st.text_input(
            "Ground truth XML folder path", value="ground_truth", key="batch_gt_folder"
        )

    with bcol2:
        batch_customer = st.selectbox(
            "Customer",
            ["smeetferrybol", "smeetferryead"],
            key="batch_customer"
        )

    with bcol3:
        max_concurrency = st.number_input(
            "Max concurrent API calls",
            min_value=1, max_value=10, value=4, step=1,
            help="How many PDFs to send to the AI API at once. Lower this "
                 "if you start seeing rate-limit or connection errors."
        )
    with bcol4:
        batch_model = st.selectbox(
            "AI Model",
            [
                "gemini-3.1-pro-preview",
                "gemini-3.5-flash",
                "gemini-2.0-flash",
                "gemini-2.5-flash",
                "gemini-2.5-pro",
                "gemini-1.5-pro",
                "gemini-1.5-flash",
            ],
            key="batch_model"
        )

    if batch_pdfs:
        st.write(f"**{len(batch_pdfs)} file(s) ready:** " + ", ".join(f.name for f in batch_pdfs))

    run_batch = st.button("▶️ Run Batch Accuracy Report")

    if run_batch:

        if not batch_pdfs:
            st.warning("Upload at least one PDF first.")
        else:

            progress_bar = st.progress(0.0)
            status_text = st.empty()

            def update_progress(done, total, current_name):
                status_text.write(f"Completed: {current_name} ({done}/{total})")
                if total:
                    progress_bar.progress(min(done / total, 1.0))

            summary_df, detail_df, error = run_batch_accuracy(
                batch_pdfs,
                gt_folder_input,
                batch_customer,
                batch_model,
                progress_callback=update_progress,
                max_concurrency=int(max_concurrency)
            )

            progress_bar.progress(1.0)

            if error:
                st.error(error)
            else:
                ok_mask = summary_df["Status"] == "OK"
                avg_accuracy = (
                    round(summary_df.loc[ok_mask, "Accuracy %"].mean(), 2)
                    if ok_mask.any() else 0.0
                )

                m1, m2 = st.columns(2)
                with m1:
                    st.metric("Average Accuracy", f"{avg_accuracy}%")
                with m2:
                    st.metric("Files Processed Successfully", f"{int(ok_mask.sum())} / {len(summary_df)}")
                # After the avg_accuracy metrics, add model-wise breakdown:
                st.subheader("Accuracy by Model")
                ok_rows = summary_df[summary_df["Status"] == "OK"]
                if not ok_rows.empty:
                    model_summary = (
                        ok_rows.groupby("Model")
                        .agg(
                            Files_Processed=("File", "count"),
                            Avg_Accuracy=("Accuracy %", "mean"),
                            Avg_Matched=("Matched Fields", "mean"),
                        )
                        .round(2)
                        .reset_index()
                        .rename(columns={
                            "Model": "Model",
                            "Files_Processed": "Files Processed",
                            "Avg_Accuracy": "Avg Accuracy %",
                            "Avg_Matched": "Avg Matched Fields"
                        })
                    )
                    st.dataframe(model_summary, use_container_width=True)
                else:
                    st.info("No successful results to group by model.")
                st.subheader("Summary")
                st.dataframe(summary_df, use_container_width=True)

                st.subheader("Field-Level Details (all files)")
                st.dataframe(detail_df, use_container_width=True)

                # Per-file breakdown so each PDF's result is easy to scan
                st.subheader("Per-File Breakdown")
                for fname in summary_df["File"]:
                    file_rows = detail_df[detail_df["File"] == fname]
                    file_row = summary_df[summary_df["File"] == fname].iloc[0]
                    label = f"{fname} — {file_row['Status']}"
                    if pd.notna(file_row["Accuracy %"]):
                        label += f" ({file_row['Accuracy %']}%)"
                    with st.expander(label):
                        if not file_rows.empty:
                            st.dataframe(
                                file_rows.drop(columns=["File"]),
                                use_container_width=True
                            )
                        else:
                            st.write("No field comparison available for this file.")

                sheets = {
                    "Summary": summary_df,
                    "Details": detail_df,
                }
                if not ok_rows.empty:
                    sheets["Model Accuracy"] = model_summary   # ← third sheet

                batch_excel = to_excel_bytes(sheets)
                st.download_button(
                    "📥 Download Batch Accuracy Report (Excel)",
                    batch_excel,
                    file_name="batch_accuracy_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
