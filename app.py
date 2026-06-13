"""
app.py  —  Streamlit web interface for the POA Generator.
Run locally:   streamlit run app.py
"""

import csv
import io
import os
import zipfile
import streamlit as st
import pandas as pd
from poa_generator import (
    FIELDS,
    generate_poa_bytes,
    validate_case,
    make_filename,
    load_cases_from_csv,
    get_csv_columns,
    get_example_row,
    get_empty_row,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="POA Generator — Tsong Law Group",
    page_icon="⚖️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Template loading
# (looks for poa_template.docx next to this script, then OneDrive fallback)
# ---------------------------------------------------------------------------

ONEDRIVE_TEMPLATE = (
    r"C:\Users\zhous\OneDrive - Tsong Law Group"
    r"\Ralph Tsong's files - Marketing\POA Pilot"
    r"\Power of Attorney - Shi & Aispuro - C&T Fertility Consultant.docx"
)
LOCAL_TEMPLATE = os.path.join(os.path.dirname(__file__), "poa_template.docx")


def load_default_template() -> bytes | None:
    """Try to load the template from local file first, then OneDrive."""
    for path in [LOCAL_TEMPLATE, ONEDRIVE_TEMPLATE]:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return f.read()
    return None


# ---------------------------------------------------------------------------
# Shared state helpers
# ---------------------------------------------------------------------------

def get_template_bytes() -> bytes | None:
    """Return template bytes from session state or default locations."""
    if "template_bytes" in st.session_state:
        return st.session_state["template_bytes"]
    default = load_default_template()
    if default:
        st.session_state["template_bytes"] = default
    return default


def group_fields_by_section() -> dict:
    sections = {}
    for f in FIELDS:
        sections.setdefault(f["section"], []).append(f)
    return sections


# ---------------------------------------------------------------------------
# UI components
# ---------------------------------------------------------------------------

def render_header():
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("POA Document Generator")
        st.caption("Tsong Law Group, A.P.C.")
    with col2:
        st.write("")
        st.write("")
        st.markdown("**⚖️ Tsong Law Group**")


def render_template_sidebar():
    with st.sidebar:
        st.header("Template")
        tpl = get_template_bytes()
        if tpl:
            st.success(f"Template loaded ({len(tpl):,} bytes)")
        else:
            st.warning("No template found. Please upload one.")

        uploaded = st.file_uploader(
            "Upload a different template (.docx)",
            type=["docx"],
            help="Upload your POA template Word document. "
                 "Field values in the template will be replaced with the case data.",
        )
        if uploaded:
            st.session_state["template_bytes"] = uploaded.read()
            st.success("Custom template loaded.")
            st.rerun()

        st.divider()
        st.caption(
            "To update the default template, save your .docx as "
            "`poa_template.docx` in the same folder as this app."
        )


def render_form_field(field: dict, prefix: str = "") -> str:
    """Render a single input field and return its value."""
    key = prefix + field["key"]
    default = field["default"]
    label = field["label"]
    if field["required"]:
        label += " *"

    value = st.text_input(
        label,
        value=st.session_state.get(key, default),
        key=key,
        help=field["help"] or None,
        placeholder=field["help"] or "",
    )
    return value.strip().upper() if field["to_upper"] else value.strip()


def collect_form_values(prefix: str = "") -> dict:
    """Read all field values from the current form state."""
    return {
        f["key"]: (
            st.session_state.get(prefix + f["key"], f["default"]).strip().upper()
            if f["to_upper"]
            else st.session_state.get(prefix + f["key"], f["default"]).strip()
        )
        for f in FIELDS
    }


# ---------------------------------------------------------------------------
# Tab 1: Single Case
# ---------------------------------------------------------------------------

def render_single_case_tab():
    st.subheader("Enter Case Information")
    st.caption("Fields marked * are required. Press Tab to move between fields.")

    sections = group_fields_by_section()

    with st.form("single_case_form"):
        cols = st.columns(2)
        col_idx = 0

        for section_name, fields in sections.items():
            with cols[col_idx % 2]:
                st.markdown(f"**{section_name}**")
                for field in fields:
                    render_form_field(field, prefix="form_")
                st.write("")
            col_idx += 1

        submitted = st.form_submit_button(
            "Generate POA Document",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        _handle_single_generate()


def _handle_single_generate():
    tpl = get_template_bytes()
    if not tpl:
        st.error("No template loaded. Upload a template in the sidebar.")
        return

    info = collect_form_values(prefix="form_")
    errors = validate_case(info)

    if errors:
        for e in errors:
            st.error(e)
        return

    with st.spinner("Generating document..."):
        try:
            doc_bytes, count = generate_poa_bytes(tpl, info)
            filename = make_filename(info)
        except Exception as exc:
            st.error(f"Error generating document: {exc}")
            return

    st.success(f"Done! {count} field replacements made.")
    st.download_button(
        label=f"Download  {filename}",
        data=doc_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
    st.info(
        "After downloading, please review:\n"
        "- Agent pronouns (her / his) if the agent's gender differs from the template\n"
        "- Notary acknowledgment state\n"
        "- Any formatting adjustments needed"
    )


# ---------------------------------------------------------------------------
# Tab 2: Batch (CSV)
# ---------------------------------------------------------------------------

def render_batch_tab():
    st.subheader("Batch Generation via CSV")

    # --- Download sample CSV ---
    with st.expander("Step 1 — Download the CSV template (fill this in Excel)", expanded=True):
        st.markdown(
            "Download the CSV template below, open it in Excel, "
            "fill in one row per case, save as CSV, then upload it here."
        )
        sample_csv = _build_sample_csv()
        st.download_button(
            "Download CSV Template",
            data=sample_csv,
            file_name="poa_cases_template.csv",
            mime="text/csv",
        )
        st.caption(
            "The first row is an example (the Shi/Aispuro case). "
            "Replace it with your real cases — one row per case."
        )

    # --- Upload CSV ---
    st.markdown("**Step 2 — Upload your filled CSV**")
    uploaded_csv = st.file_uploader(
        "Upload CSV file",
        type=["csv"],
        label_visibility="collapsed",
    )

    if not uploaded_csv:
        return

    csv_bytes = uploaded_csv.read()
    cases, errors = load_cases_from_csv(csv_bytes)

    if errors:
        st.warning(f"{len(errors)} row(s) had issues and were skipped:")
        for row_num, msg in errors:
            st.caption(f"  Row {row_num}: {msg}")

    if not cases:
        st.error("No valid cases found in the CSV.")
        return

    st.success(f"{len(cases)} case(s) ready to generate.")

    # Preview table
    preview_df = pd.DataFrame(cases)
    display_cols = ["principal_name", "principal_passport", "surrogate_name", "agent_name", "due_date"]
    st.dataframe(
        preview_df[[c for c in display_cols if c in preview_df.columns]],
        use_container_width=True,
    )

    tpl = get_template_bytes()
    if not tpl:
        st.error("No template loaded. Upload a template in the sidebar.")
        return

    if st.button("Generate All Documents", type="primary", use_container_width=True):
        _handle_batch_generate(tpl, cases)


def _handle_batch_generate(tpl: bytes, cases: list):
    progress = st.progress(0, text="Generating documents...")
    zip_buf = io.BytesIO()
    failed = []

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, info in enumerate(cases):
            try:
                doc_bytes, _ = generate_poa_bytes(tpl, info)
                filename = make_filename(info)
                zf.writestr(filename, doc_bytes)
            except Exception as exc:
                failed.append((i + 1, str(exc)))
            progress.progress((i + 1) / len(cases), text=f"Generating {i + 1}/{len(cases)}...")

    progress.empty()

    if failed:
        for row, msg in failed:
            st.warning(f"Case {row} failed: {msg}")

    count = len(cases) - len(failed)
    st.success(f"Generated {count} document(s).")
    st.download_button(
        label=f"Download All ({count} documents as ZIP)",
        data=zip_buf.getvalue(),
        file_name="POA_Documents.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )


def _build_sample_csv() -> bytes:
    """Build a CSV file with column headers + one example row."""
    columns = get_csv_columns()
    example = get_example_row()

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=columns, lineterminator="\n")
    writer.writeheader()
    writer.writerow(example)
    return buf.getvalue().encode("utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    render_header()
    render_template_sidebar()
    st.divider()

    tab_single, tab_batch = st.tabs(["Single Case", "Batch (CSV)"])

    with tab_single:
        render_single_case_tab()

    with tab_batch:
        render_batch_tab()


if __name__ == "__main__":
    main()
