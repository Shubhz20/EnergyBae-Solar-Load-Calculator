"""
Energybae - MSEDCL Bill Automation
Streamlit UI: upload bill -> Gemini extract -> verify -> fill Excel template -> download.
"""

import os
import io
import json
import tempfile
from pathlib import Path

import streamlit as st

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from utils.extractor import extract_bill_data
from utils.excel_handler import fill_solar_load_template

st.set_page_config(
    page_title="Energybae | MSEDCL Bill Automation",
    page_icon="bolt",
    layout="centered",
)

st.title("Energybae - MSEDCL Bill Automation")
st.caption(
    "Upload an MSEDCL electricity bill (PDF/Image). The tool extracts key fields "
    "via Gemini 1.5 Flash, lets you verify them, and fills the Solar Load Excel "
    "Template without touching any formulas."
)

# ---------------------------------------------------------------------------
# API key - resolved silently from Streamlit secrets or env. End users never
# see it. If it's missing, that's an admin/deployment problem, not an end-user
# problem - we surface a clean error instead of an input field.
# ---------------------------------------------------------------------------
def _resolve_api_key() -> str:
    try:
        if "GOOGLE_API_KEY" in st.secrets:
            return str(st.secrets["GOOGLE_API_KEY"]).strip()
    except Exception:
        pass
    return os.getenv("GOOGLE_API_KEY", "").strip()


api_key = _resolve_api_key()

# ---------------------------------------------------------------------------
# Bundled Solar Load template - shipped with the repo so the app always has
# something to fill. Sales team can also upload their own.
# ---------------------------------------------------------------------------
BUNDLED_TEMPLATE_PATH = Path(__file__).parent / "templates" / "solar_load_template.xlsx"

with st.sidebar:
    st.header("Template")
    uploaded_template = st.file_uploader(
        "Use a different template (optional)",
        type=["xlsx"],
        help="Leave empty to use the bundled Solar Load template.",
    )
    if uploaded_template is None and BUNDLED_TEMPLATE_PATH.exists():
        st.caption("Using bundled Solar Load template.")

# ---------------------------------------------------------------------------
# Main - bill upload
# ---------------------------------------------------------------------------
bill_file = st.file_uploader(
    "Upload MSEDCL Bill",
    type=["pdf", "png", "jpg", "jpeg"],
)

if "extracted" not in st.session_state:
    st.session_state.extracted = None

# ---------------------------------------------------------------------------
# Step 1 - Extract
# ---------------------------------------------------------------------------
if bill_file is not None:
    if st.button("Extract Data with Gemini", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please provide a Gemini API key in the sidebar.")
        else:
            with st.spinner("Calling Gemini 1.5 Flash..."):
                try:
                    suffix = Path(bill_file.name).suffix.lower()
                    with tempfile.NamedTemporaryFile(
                        delete=False, suffix=suffix
                    ) as tmp:
                        tmp.write(bill_file.getvalue())
                        tmp_path = tmp.name

                    data = extract_bill_data(tmp_path, api_key=api_key)
                    st.session_state.extracted = data
                    st.success("Extraction complete. Verify the fields below.")
                except Exception as e:
                    st.error(f"Extraction failed: {e}")

# ---------------------------------------------------------------------------
# Step 2 - Verify (Human-in-the-Loop)
# ---------------------------------------------------------------------------
if st.session_state.extracted:
    st.subheader("Verify Extracted Fields")
    st.caption(
        "Edit any value before generating the Excel. AI is fast, but you are "
        "the source of truth on financial billing."
    )

    data = st.session_state.extracted

    col1, col2 = st.columns(2)
    with col1:
        consumer_name = st.text_input(
            "Consumer Name", value=str(data.get("consumer_name") or "")
        )
        consumer_number = st.text_input(
            "Consumer Number", value=str(data.get("consumer_number") or "")
        )
        billing_unit = st.text_input(
            "Billing Unit", value=str(data.get("billing_unit") or "")
        )
    with col2:
        connected_load_kw = st.text_input(
            "Connected Load (kW)",
            value=str(data.get("connected_load_kw") or ""),
        )
        tariff_category = st.text_input(
            "Tariff Category", value=str(data.get("tariff_category") or "")
        )
        avg_monthly_consumption = st.text_input(
            "Avg Monthly Consumption (kWh)",
            value=str(data.get("avg_monthly_consumption") or ""),
        )

    verified = {
        "consumer_name": consumer_name,
        "consumer_number": consumer_number,
        "billing_unit": billing_unit,
        "connected_load_kw": connected_load_kw,
        "tariff_category": tariff_category,
        "avg_monthly_consumption": avg_monthly_consumption,
    }

    with st.expander("Show raw JSON from Gemini"):
        st.json(data)

    # -----------------------------------------------------------------------
    # Step 3 - Fill template & download
    # -----------------------------------------------------------------------
    st.subheader("Generate Filled Excel")

    if uploaded_template is not None:
        template_bytes = uploaded_template.getvalue()
        template_source = "uploaded template"
    elif BUNDLED_TEMPLATE_PATH.exists():
        template_bytes = BUNDLED_TEMPLATE_PATH.read_bytes()
        template_source = "bundled template"
    else:
        template_bytes = None
        template_source = None

    if template_bytes is None:
        st.error(
            "No template available. Upload one in the sidebar or contact "
            "the admin to add the bundled Solar Load template."
        )
    else:
        if st.button("Fill Template & Prepare Download", use_container_width=True):
            try:
                output_bytes = fill_solar_load_template(template_bytes, verified)

                out_name = (
                    f"SolarLoad_{(consumer_number or 'output').strip().replace(' ', '_')}.xlsx"
                )
                st.success(f"Excel generated using the {template_source}. Formulas preserved.")
                st.download_button(
                    label="Download Filled Excel",
                    data=output_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Failed to write template: {e}")
