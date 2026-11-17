import hashlib
import io
import zipfile
from datetime import datetime

import streamlit as st

from sales.southcomp_engine import (
    build_output_filename,
    describe_input_kind,
    generate_southcomp_quote,
)


def render_southcomp_tool(team, update_usage) -> None:
    st.title("💼 Dell Quotation Southcomp Polaris")

    uploaded_files = st.file_uploader(
        "Upload Dell BOQ Excel, PDF or Word (.docx)",
        type=["xlsx", "xlsm", "xls", "pdf", "docx"],
        accept_multiple_files=True,
        key="southcomp_uploader",
    )

    margin_percent = st.number_input(
        "Default Margin %",
        min_value=0.0,
        max_value=99.5,
        value=5.0,
        step=0.5,
        key="southcomp_margin",
    )

    exchange_rate = st.number_input(
        "Exchange Rate (EUR → USD)",
        min_value=0.0,
        value=0.92,
        step=0.01,
        format="%.4f",
        key="southcomp_eur_exchange_rate",
    )

    # Session state initialisation
    for key, default in [
        ("southcomp_outputs", []),
        ("southcomp_uploaded_hash", None),
        ("southcomp_uploaded_inputs", []),
        ("southcomp_last_margin_percent", None),
        ("southcomp_last_exchange_rate", None),
        ("southcomp_generation_success", False),
        ("southcomp_last_error", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    def reset_outputs():
        st.session_state["southcomp_outputs"] = []
        st.session_state["southcomp_generation_success"] = False
        st.session_state["southcomp_last_error"] = None

    # Detect file change and reset outputs
    if uploaded_files:
        inputs = [(f.name, f.getvalue()) for f in uploaded_files]
        combined = hashlib.sha256()
        for name, data in inputs:
            combined.update(name.encode("utf-8"))
            combined.update(data)
        uploaded_hash = combined.hexdigest()
        if st.session_state["southcomp_uploaded_hash"] != uploaded_hash:
            reset_outputs()
            st.session_state["southcomp_uploaded_hash"] = uploaded_hash
            st.session_state["southcomp_uploaded_inputs"] = inputs
    else:
        inputs = st.session_state.get("southcomp_uploaded_inputs", [])

    # Reset on parameter change
    if (
        st.session_state.get("southcomp_last_margin_percent") not in (None, margin_percent)
        or st.session_state.get("southcomp_last_exchange_rate") not in (None, exchange_rate)
    ):
        reset_outputs()

    col1, _ = st.columns([1, 1])
    with col1:
        generate_clicked = st.button(
            "🚀 Generate Quotation",
            key="southcomp_generate_btn",
            use_container_width=True,
        )

    if st.session_state.get("southcomp_generation_success", False):
        st.success("✅ Quotation generated successfully!")
        st.session_state["southcomp_generation_success"] = False

    if generate_clicked:
        if not inputs:
            st.warning("Please upload at least one file first.")
        else:
            try:
                st.session_state["southcomp_last_error"] = None
                outputs = []
                with st.spinner("⚙️ Generating quotations..."):
                    for source_name, input_bytes in inputs:
                        entry = {
                            "source_name": source_name,
                            "input_kind": describe_input_kind(input_bytes),
                        }
                        for target_currency in ("EUR", "USD"):
                            entry[target_currency] = generate_southcomp_quote(
                                input_bytes=input_bytes,
                                margin_percent=margin_percent,
                                currency_code=target_currency,
                                exchange_rate=exchange_rate if target_currency == "EUR" else 1.0,
                            )
                            entry[f"{target_currency}_name"] = build_output_filename(
                                target_currency, source_name
                            )
                        outputs.append(entry)

                    st.session_state["southcomp_outputs"] = outputs
                    st.session_state["southcomp_last_margin_percent"] = margin_percent
                    st.session_state["southcomp_last_exchange_rate"] = exchange_rate
                    st.session_state["southcomp_generation_success"] = True

                st.success("✅ Quotation generated successfully.")

            except Exception as e:
                st.session_state["southcomp_last_error"] = str(e)
                st.error(str(e))
                st.exception(e)

    outputs = st.session_state.get("southcomp_outputs", [])

    if outputs:
        st.markdown("### Download your files")

        def usage_counts(entries):
            pdf_count = sum(1 for o in entries if o["input_kind"] == "pdf")
            excel_count = sum(
                1 for o in entries if o["input_kind"] in ("qar", "boq_grouped", "boq_generic")
            )
            return pdf_count, excel_count

        if len(outputs) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for entry in outputs:
                    for currency in ("EUR", "USD"):
                        zip_file.writestr(entry[f"{currency}_name"], entry[currency])
            zip_buffer.seek(0)
            all_pdf_count, all_excel_count = usage_counts(outputs)
            st.download_button(
                label="⬇️ Download all quotations (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"Southcomp_Polaris_quotations_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                key="southcomp_download_zip",
                on_click=lambda: update_usage(
                    "southcomp polaris-multi",
                    team,
                    pdf_count=all_pdf_count,
                    excel_count=all_excel_count,
                ),
                use_container_width=True,
            )

        for idx, entry in enumerate(outputs):
            if len(outputs) > 1:
                st.markdown(f"**📄 {entry['source_name']}**")
            input_kind = entry["input_kind"]
            pdf_count, excel_count = usage_counts([entry])
            eur_col, usd_col = st.columns(2)
            with eur_col:
                st.download_button(
                    label="⬇️ Download EUR quotation",
                    data=entry["EUR"],
                    file_name=entry["EUR_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"southcomp_download_eur_{idx}",
                    on_click=lambda k=input_kind, p=pdf_count, x=excel_count: update_usage(
                        f"southcomp polaris-{k}-EUR", team, pdf_count=p, excel_count=x
                    ),
                    use_container_width=True,
                )
            with usd_col:
                st.download_button(
                    label="⬇️ Download USD quotation",
                    data=entry["USD"],
                    file_name=entry["USD_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"southcomp_download_usd_{idx}",
                    on_click=lambda k=input_kind, p=pdf_count, x=excel_count: update_usage(
                        f"southcomp polaris-{k}-USD", team, pdf_count=p, excel_count=x
                    ),
                    use_container_width=True,
                )

    if not uploaded_files and not outputs:
        st.info("Upload one or more Dell BOQ Excel, PDF or Word (.docx) files, then click Generate Quotation.")
