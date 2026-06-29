import hashlib

import streamlit as st

from southcomp_engine import (
    build_output_filename,
    detect_template_type,
    generate_southcomp_quote,
)


def render_southcomp_tool() -> None:
    st.title("💼 Dell Quotation Southcomp Polaris")

    uploaded = st.file_uploader(
        "Upload Dell BOQ Excel or PDF",
        type=["xlsx", "xlsm", "xls", "pdf"],
        key="southcomp_uploader",
    )

    margin_percent = st.number_input(
        "Default Margin %",
        min_value=0.0,
        max_value=100.0,
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
        ("southcomp_output_bytes_eur", None),
        ("southcomp_output_bytes_usd", None),
        ("southcomp_output_name_eur", None),
        ("southcomp_output_name_usd", None),
        ("southcomp_uploaded_hash", None),
        ("southcomp_uploaded_bytes", None),
        ("southcomp_last_uploaded_name", None),
        ("southcomp_last_margin_percent", None),
        ("southcomp_last_exchange_rate", None),
        ("southcomp_generation_success", False),
        ("southcomp_last_error", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # Detect file change and reset outputs
    if uploaded is not None:
        uploaded_bytes = uploaded.getvalue()
        uploaded_hash = hashlib.sha256(uploaded_bytes).hexdigest()
        if (
            st.session_state["southcomp_last_uploaded_name"] != uploaded.name
            or st.session_state["southcomp_uploaded_hash"] != uploaded_hash
        ):
            st.session_state["southcomp_output_bytes_eur"] = None
            st.session_state["southcomp_output_bytes_usd"] = None
            st.session_state["southcomp_output_name_eur"] = None
            st.session_state["southcomp_output_name_usd"] = None
            st.session_state["southcomp_generation_success"] = False
            st.session_state["southcomp_last_error"] = None
            st.session_state["southcomp_uploaded_hash"] = uploaded_hash
            st.session_state["southcomp_uploaded_bytes"] = uploaded_bytes
            st.session_state["southcomp_last_uploaded_name"] = uploaded.name
    else:
        uploaded_bytes = st.session_state.get("southcomp_uploaded_bytes")

    # Reset on parameter change
    if (
        st.session_state.get("southcomp_last_margin_percent") not in (None, margin_percent)
        or st.session_state.get("southcomp_last_exchange_rate") not in (None, exchange_rate)
    ):
        st.session_state["southcomp_output_bytes_eur"] = None
        st.session_state["southcomp_output_bytes_usd"] = None
        st.session_state["southcomp_output_name_eur"] = None
        st.session_state["southcomp_output_name_usd"] = None
        st.session_state["southcomp_generation_success"] = False
        st.session_state["southcomp_last_error"] = None

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
        if uploaded is None and st.session_state.get("southcomp_uploaded_bytes") is None:
            st.warning("Please upload a file first.")
        else:
            try:
                st.session_state["southcomp_last_error"] = None
                input_bytes = st.session_state.get("southcomp_uploaded_bytes") or uploaded.getvalue()
                if input_bytes is None:
                    raise ValueError("Uploaded file bytes are missing.")

                with st.spinner("⚙️ Generating quotation..."):
                    generated_outputs: dict[str, bytes] = {}
                    for target_currency in ("EUR", "USD"):
                        generated_outputs[target_currency] = generate_southcomp_quote(
                            input_bytes=input_bytes,
                            margin_percent=margin_percent,
                            currency_code=target_currency,
                            exchange_rate=exchange_rate if target_currency == "EUR" else 1.0,
                        )

                    st.session_state["southcomp_output_bytes_eur"] = generated_outputs["EUR"]
                    st.session_state["southcomp_output_bytes_usd"] = generated_outputs["USD"]
                    st.session_state["southcomp_output_name_eur"] = build_output_filename("EUR")
                    st.session_state["southcomp_output_name_usd"] = build_output_filename("USD")
                    st.session_state["southcomp_last_margin_percent"] = margin_percent
                    st.session_state["southcomp_last_exchange_rate"] = exchange_rate
                    st.session_state["southcomp_generation_success"] = True

                st.success("✅ Quotation generated successfully.")

            except Exception as e:
                st.session_state["southcomp_last_error"] = str(e)
                st.error(str(e))
                st.exception(e)

    eur_bytes = st.session_state.get("southcomp_output_bytes_eur")
    usd_bytes = st.session_state.get("southcomp_output_bytes_usd")

    if eur_bytes or usd_bytes:
        st.markdown("### Download your files")
        if eur_bytes:
            st.download_button(
                label="⬇️ Download EUR quotation",
                data=eur_bytes,
                file_name=st.session_state.get("southcomp_output_name_eur", "quotation_eur.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="southcomp_download_eur",
                use_container_width=True,
            )
        if usd_bytes:
            st.download_button(
                label="⬇️ Download USD quotation",
                data=usd_bytes,
                file_name=st.session_state.get("southcomp_output_name_usd", "quotation_usd.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="southcomp_download_usd",
                use_container_width=True,
            )

    if uploaded is None and not eur_bytes and not usd_bytes:
        st.info("Upload Dell BOQ Excel or PDF, then click Generate Quotation.")
