import hashlib
from datetime import datetime

import streamlit as st

from sales.southcomp_engine import (
    build_item_creation_filename,
    generate_item_creation_excel_from_inputs,
)


def render_southcomp_item_creation_tool(team, update_usage) -> None:
    st.title("💼 Dell Quotation Southcomp Polaris — Item Creation")
    st.write(
        "Upload one or more Dell quotes (Excel BOQ, PDF or Word) and download an "
        "Item / Description list built from each item's Base, Display, Processor, "
        "Memory, Storage, Wireless, Operating System and Primary Battery configuration."
    )

    uploaded_files = st.file_uploader(
        "Upload Dell BOQ Excel, PDF or Word (.docx)",
        type=["xlsx", "xlsm", "xls", "pdf", "docx"],
        accept_multiple_files=True,
        key="southcomp_items_uploader",
    )

    for key, default in [
        ("southcomp_items_output_bytes", None),
        ("southcomp_items_output_name", None),
        ("southcomp_items_row_count", 0),
        ("southcomp_items_uploaded_hash", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    if uploaded_files:
        inputs = [(f.name, f.getvalue()) for f in uploaded_files]
        combined = hashlib.sha256()
        for name, data in inputs:
            combined.update(name.encode("utf-8"))
            combined.update(data)
        uploaded_hash = combined.hexdigest()
        if st.session_state["southcomp_items_uploaded_hash"] != uploaded_hash:
            st.session_state["southcomp_items_uploaded_hash"] = uploaded_hash
            st.session_state["southcomp_items_output_bytes"] = None
            st.session_state["southcomp_items_output_name"] = None
            st.session_state["southcomp_items_row_count"] = 0
    else:
        inputs = []

    if st.button("🚀 Generate Item List", key="southcomp_items_generate_btn", use_container_width=True):
        if not inputs:
            st.warning("Please upload at least one file first.")
        else:
            try:
                with st.spinner("⚙️ Extracting item configurations..."):
                    xlsx_bytes, row_count = generate_item_creation_excel_from_inputs(inputs)
                    st.session_state["southcomp_items_output_bytes"] = xlsx_bytes
                    st.session_state["southcomp_items_output_name"] = build_item_creation_filename()
                    st.session_state["southcomp_items_row_count"] = row_count
                if row_count:
                    st.success(f"✅ Item list generated — {row_count} item(s) extracted.")
                else:
                    st.warning(
                        "No items with a recognizable configuration were found in the "
                        "uploaded file(s). This can happen for quotes with no per-item "
                        "Product Details / Configuration breakdown (e.g. an extended "
                        "warranty quote)."
                    )
            except Exception as e:
                st.session_state["southcomp_items_output_bytes"] = None
                st.session_state["southcomp_items_output_name"] = None
                st.error(str(e))

    if st.session_state.get("southcomp_items_output_bytes"):
        pdf_count = sum(1 for name, _ in inputs if name.lower().endswith(".pdf")) if inputs else 0
        excel_count = sum(
            1 for name, _ in inputs if name.lower().endswith((".xlsx", ".xlsm", ".xls", ".docx"))
        ) if inputs else 0
        st.download_button(
            label="⬇️ Download item list (Excel)",
            data=st.session_state["southcomp_items_output_bytes"],
            file_name=st.session_state.get("southcomp_items_output_name") or "Southcomp_Polaris_Item_Creation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="southcomp_items_download_btn",
            use_container_width=True,
            on_click=lambda: update_usage(
                "southcomp polaris-item-creation",
                team,
                pdf_count=pdf_count,
                excel_count=excel_count,
            ),
        )

    if not uploaded_files and not st.session_state.get("southcomp_items_output_bytes"):
        st.info("Upload one or more Dell quotes, then click Generate Item List.")
