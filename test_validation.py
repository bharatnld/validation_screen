# import streamlit as st
# import pandas as pd
# import asyncio
# import json
# from api_client import process_file

# st.set_page_config(
#     page_title="OCR Validation Portal",
#     layout="wide"
# )

# # ===========================
# # CSS
# # ===========================

# st.markdown("""
# <style>

# .block-container{
#     padding-top:0.5rem;
#     max-width:100%;
# }

# .section-box{
#     border:1px solid #dcdcdc;
#     padding:10px;
#     border-radius:8px;
#     margin-bottom:10px;
# }

# .order-header{
#     font-size:28px;
#     font-weight:bold;
#     color:#2d89ef;
# }

# </style>
# """, unsafe_allow_html=True)

# # ===========================
# # HELPERS
# # ===========================

# def get_extracted_data(response):

#     return (
#         response
#         .get("data", {})
#         .get("extracted_data", {})
#         .get("gpt_extraction_output", {})
#     )

# # ===========================
# # HEADER
# # ===========================

# st.markdown(
#     '<div class="order-header">OCR Validation Portal</div>',
#     unsafe_allow_html=True
# )

# st.divider()

# # ===========================
# # TOP BAR
# # ===========================

# col1, col2, col3 = st.columns([3,1,1])

# with col1:
#     uploaded_file = st.file_uploader(
#         "Upload PDF",
#         type=["pdf"]
#     )

# with col2:
#     customer_name = st.selectbox(
#         "Customer",
#         [
#             "smeetferrybol",
#             "smeetferryead"
#         ]
#     )

# with col3:
#     process_btn = st.button(
#         "🚀 Process"
#     )

# # ===========================
# # PROCESS FILE
# # ===========================

# if process_btn and uploaded_file:

#     with st.spinner("Running OCR..."):

#         response = asyncio.run(
#             process_file(
#                 uploaded_file,
#                 customer_name
#             )
#         )

#         st.session_state["api_response"] = response

# # ===========================
# # DISPLAY DATA
# # ===========================

# if "api_response" in st.session_state:

#     response = st.session_state["api_response"]

#     result = get_extracted_data(response)

#     shipment = result.get(
#         "shipment_metadata",
#         {}
#     )

#     exporter = result.get(
#         "exporter_information",
#         {}
#     )

#     importer = result.get(
#         "importer_information",
#         {}
#     )

#     items = result.get(
#         "items",
#         []
#     )

#     left, right = st.columns([1,2])

#     # ==================================
#     # LEFT PANEL
#     # ==================================

#     with left:

#         with st.expander(
#             "SHIPMENT METADATA",
#             expanded=True
#         ):

#             shipment["customtype"] = st.text_input(
#                 "Custom Type",
#                 shipment.get("customtype") or ""
#             )

#             shipment["MRN"] = st.text_input(
#                 "MRN",
#                 shipment.get("MRN") or ""
#             )

#             shipment["total_quantity"] = st.text_input(
#                 "Total Quantity",
#                 shipment.get("total_quantity") or ""
#             )

#             shipment["total_gross_weight"] = st.text_input(
#                 "Gross Weight",
#                 shipment.get("total_gross_weight") or ""
#             )

#             shipment["date"] = st.text_input(
#                 "Date",
#                 shipment.get("date") or ""
#             )

#         with st.expander(
#             "EXPORTER",
#             expanded=True
#         ):

#             exporter["name"] = st.text_input(
#                 "Exporter Name",
#                 exporter.get("name") or ""
#             )

#             exporter["tax_id"] = st.text_input(
#                 "Exporter Tax ID",
#                 exporter.get("tax_id") or ""
#             )

#             address = exporter.get(
#                 "address",
#                 {}
#             )

#             address["street_and_nr"] = st.text_input(
#                 "Street",
#                 address.get("street_and_nr") or ""
#             )

#             address["city"] = st.text_input(
#                 "City",
#                 address.get("city") or ""
#             )

#             address["country"] = st.text_input(
#                 "Country",
#                 address.get("country") or ""
#             )

#             address["postal_code"] = st.text_input(
#                 "Postal Code",
#                 address.get("postal_code") or ""
#             )

#             exporter["address"] = address

#         with st.expander(
#             "IMPORTER",
#             expanded=True
#         ):

#             importer["name"] = st.text_input(
#                 "Importer Name",
#                 importer.get("name") or ""
#             )

#             importer["tax_id"] = st.text_input(
#                 "Importer Tax ID",
#                 importer.get("tax_id") or ""
#             )

#             address = importer.get(
#                 "address",
#                 {}
#             )

#             address["street_and_nr"] = st.text_input(
#                 "Street ",
#                 address.get("street_and_nr") or ""
#             )

#             address["city"] = st.text_input(
#                 "City ",
#                 address.get("city") or ""
#             )

#             address["country"] = st.text_input(
#                 "Country ",
#                 address.get("country") or ""
#             )

#             address["postal_code"] = st.text_input(
#                 "Postal Code ",
#                 address.get("postal_code") or ""
#             )

#             importer["address"] = address

#     # ==================================
#     # PDF PREVIEW
#     # ==================================

#     with right:

#         st.subheader("PDF Preview")

#         if uploaded_file:

#             pdf_bytes = uploaded_file.read()

#             st.download_button(
#                 "Download PDF",
#                 pdf_bytes,
#                 file_name=uploaded_file.name
#             )

#             st.info(
#                 "Use streamlit-pdf-viewer package for embedded PDF view."
#             )

#             st.json(response)

#     # ==================================
#     # ITEMS GRID
#     # ==================================

#     st.divider()

#     st.subheader("LINE ITEMS")

#     if items:

#         df = pd.DataFrame(items)

#         edited_df = st.data_editor(
#             df,
#             use_container_width=True,
#             num_rows="dynamic",
#             hide_index=True
#         )

#         result["items"] = edited_df.to_dict(
#             orient="records"
#         )

#     # ==================================
#     # SAVE RESULT
#     # ==================================

#     result["shipment_metadata"] = shipment
#     result["exporter_information"] = exporter
#     result["importer_information"] = importer

#     st.divider()

#     c1, c2, c3 = st.columns(3)

#     with c1:

#         if st.button("✅ Approve"):

#             with open(
#                 "validated_output.json",
#                 "w",
#                 encoding="utf-8"
#             ) as f:

#                 json.dump(
#                     result,
#                     f,
#                     indent=4,
#                     ensure_ascii=False
#                 )

#             st.success(
#                 "Validation Approved"
#             )

#     with c2:

#         if st.button("❌ Reject"):

#             st.error(
#                 "Validation Rejected"
#             )

#     with c3:

#         st.download_button(
#             "📥 Download JSON",
#             json.dumps(
#                 result,
#                 indent=4,
#                 ensure_ascii=False
#             ),
#             file_name="validated_output.json",
#             mime="application/json"
#         )




import streamlit as st
import pandas as pd
import asyncio
import json
from streamlit_pdf_viewer import pdf_viewer          # ← added
from api_client import process_file               # ← keep your original import

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

def get_extracted_data(response):
    return (
        response
        .get("data", {})
        .get("extracted_data", {})
        .get("gpt_extraction_output", {})
    )

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

col1, col2, col3 = st.columns([3, 1, 1])

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
            "smeetferryead"
        ]
    )

with col3:
    process_btn = st.button("🚀 Process")

# ===========================
# PROCESS FILE
# ===========================

if process_btn and uploaded_file:

    with st.spinner("Running OCR..."):

        response = asyncio.run(process_file(uploaded_file, customer_name))

        st.session_state["api_response"] = response

# ===========================
# DISPLAY DATA
# ===========================

if "api_response" in st.session_state:

    response = st.session_state["api_response"]
    result   = get_extracted_data(response)

    shipment = result.get("shipment_metadata",   {})
    exporter = result.get("exporter_information", {})
    importer = result.get("importer_information", {})
    items    = result.get("items", [])

    left, right = st.columns([1, 2])

    # ==================================
    # LEFT PANEL
    # ==================================

    with left:

        with st.expander("SHIPMENT METADATA", expanded=True):

            shipment["customtype"] = st.text_input(
                "Custom Type", shipment.get("customtype") or "")

            shipment["MRN"] = st.text_input(
                "MRN", shipment.get("MRN") or "")

            shipment["total_quantity"] = st.text_input(
                "Total Quantity", shipment.get("total_quantity") or "")

            shipment["total_gross_weight"] = st.text_input(
                "Gross Weight", shipment.get("total_gross_weight") or "")

            shipment["date"] = st.text_input(
                "Date", shipment.get("date") or "")

            shipment["destination_office_text"] = st.text_input(
                "Destination Office Text", shipment.get("destination_office_text") or "")

            shipment["destination_office_country"] = st.text_input(
                "Destination Office Country", shipment.get("destination_office_country") or "")

            shipment["shipment_reference"] = st.text_input(
                "Shipment Reference", shipment.get("shipment_reference") or "")


        with st.expander("EXPORTER", expanded=True):

            exporter["name"] = st.text_input(
                "Exporter Name", exporter.get("name") or "")

            exporter["tax_id"] = st.text_input(
                "Exporter Tax ID", exporter.get("tax_id") or "")

            address = exporter.get("address", {})

            address["street_and_nr"] = st.text_input(
                "Street", address.get("street_and_nr") or "")
            address["city"]          = st.text_input(
                "City",   address.get("city") or "")
            address["country"]       = st.text_input(
                "Country", address.get("country") or "")
            address["postal_code"]   = st.text_input(
                "Postal Code", address.get("postal_code") or "")

            exporter["address"] = address

        with st.expander("IMPORTER", expanded=True):

            importer["name"] = st.text_input(
                "Importer Name", importer.get("name") or "")

            importer["tax_id"] = st.text_input(
                "Importer Tax ID", importer.get("tax_id") or "")

            address = importer.get("address", {})

            address["street_and_nr"] = st.text_input(
                "Street ", address.get("street_and_nr") or "")
            address["city"]          = st.text_input(
                "City ",   address.get("city") or "")
            address["country"]       = st.text_input(
                "Country ", address.get("country") or "")
            address["postal_code"]   = st.text_input(
                "Postal Code ", address.get("postal_code") or "")

            importer["address"] = address

    # ==================================
    # PDF PREVIEW  ← fixed
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
