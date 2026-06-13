"""
app.py  —  Streamlit web interface for the POA Generator.
Run locally:   streamlit run app.py
"""

import csv
import io
import os
import zipfile
from typing import Optional

import pandas as pd
import streamlit as st

from poa_generator import (
    FIELDS,
    generate_poa_bytes,
    validate_case,
    make_filename,
    load_cases_from_csv,
    get_csv_columns,
    get_example_row,
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
# ---------------------------------------------------------------------------

ONEDRIVE_TEMPLATE = (
    r"C:\Users\zhous\OneDrive - Tsong Law Group"
    r"\Ralph Tsong's files - Marketing\POA Pilot"
    r"\Power of Attorney - Shi & Aispuro - C&T Fertility Consultant.docx"
)
LOCAL_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "poa_template.docx")


def load_default_template() -> Optional[bytes]:
    for path in [LOCAL_TEMPLATE, ONEDRIVE_TEMPLATE]:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return f.read()
    return None


def get_template_bytes() -> Optional[bytes]:
    if "template_bytes" in st.session_state:
        return st.session_state["template_bytes"]
    tpl = load_default_template()
    if tpl:
        st.session_state["template_bytes"] = tpl
    return tpl


def group_fields_by_section() -> dict:
    sections = {}
    for f in FIELDS:
        sections.setdefault(f["section"], []).append(f)
    return sections


# ---------------------------------------------------------------------------
# Sidebar — template management
# ---------------------------------------------------------------------------

def render_sidebar():
    with st.sidebar:
        st.header("Template")

        tpl = get_template_bytes()
        if tpl:
            st.success(f"Template ready ({len(tpl):,} bytes)")
        else:
            st.error("No template loaded — upload one below before generating.")

        uploaded = st.file_uploader("Upload POA template (.docx)", type=["docx"])
        if uploaded and "template_bytes" not in st.session_state:
            st.session_state["template_bytes"] = uploaded.read()
            st.session_state.pop("single_result", None)

        st.divider()
        st.caption(
            "Upload the file:  \n"
            "`Power of Attorney - Shi & Aispuro...docx`  \n"
            "from the POA Pilot folder on OneDrive."
        )


# ---------------------------------------------------------------------------
# Tab 1 — Single case
# ---------------------------------------------------------------------------

def render_single_tab():
    # Show template warning prominently if missing
    if not get_template_bytes():
        st.warning("Upload the POA template in the sidebar before generating a document.")
        return

    st.subheader("Enter Case Information")
    st.caption("Fields marked * are required.")

    sections = group_fields_by_section()

    with st.form("poa_form", clear_on_submit=False):
        left, right = st.columns(2)
        for i, (section_name, fields) in enumerate(sections.items()):
            col = left if i % 2 == 0 else right
            with col:
                st.markdown(f"**{section_name}**")
                for field in fields:
                    label = field["label"] + (" *" if field["required"] else "")
                    st.text_input(
                        label,
                        value=field["default"],
                        key=f"f_{field['key']}",
                        help=field["help"] or None,
                        placeholder=field["help"] or "",
                    )
                st.write("")

        st.divider()
        st.markdown("**ID / Passport Photos**")
        photo_col1, photo_col2 = st.columns(2)
        with photo_col1:
            st.file_uploader(
                "Principal's Passport / ID",
                type=["jpg", "jpeg", "png"],
                key="photo_principal",
                help="Photo of the intended parent's passport or government ID",
            )
        with photo_col2:
            st.file_uploader(
                "Agent's Passport / ID",
                type=["jpg", "jpeg", "png"],
                key="photo_agent",
                help="Photo of the attorney-in-fact's passport or government ID",
            )

        submitted = st.form_submit_button(
            "Generate POA Document",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        _run_single_generate()

    # Always show the download button if a result exists in session state
    _render_single_result()


def _run_single_generate():
    """Validate, generate, and store result in session state."""
    tpl = get_template_bytes()
    if not tpl:
        st.session_state["single_result"] = {"error": "No template loaded."}
        return

    # Read field values from session state (set by the form widgets above)
    info = {}
    for field in FIELDS:
        raw = st.session_state.get(f"f_{field['key']}", field["default"]) or ""
        info[field["key"]] = raw.strip().upper() if field["to_upper"] else raw.strip()

    errors = validate_case(info)
    if errors:
        st.session_state["single_result"] = {"errors": errors}
        return

    # Read uploaded photos (file_uploader returns None if nothing uploaded)
    principal_photo = None
    agent_photo = None
    p = st.session_state.get("photo_principal")
    a = st.session_state.get("photo_agent")
    if p is not None:
        principal_photo = p.read() if hasattr(p, "read") else p
    if a is not None:
        agent_photo = a.read() if hasattr(a, "read") else a

    try:
        doc_bytes, count = generate_poa_bytes(tpl, info, principal_photo, agent_photo)
        filename = make_filename(info)
        st.session_state["single_result"] = {
            "doc_bytes": doc_bytes,
            "filename": filename,
            "count": count,
        }
    except Exception as exc:
        st.session_state["single_result"] = {"error": str(exc)}


def _render_single_result():
    """Render errors or download button from session state."""
    result = st.session_state.get("single_result")
    if not result:
        return

    if "error" in result:
        st.error(result["error"])
    elif "errors" in result:
        for e in result["errors"]:
            st.error(e)
    else:
        st.success(f"Document ready — {result['count']} fields replaced.")
        st.download_button(
            label=f"Download: {result['filename']}",
            data=result["doc_bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
        st.caption(
            "After downloading, please check: agent pronouns (her/his), "
            "notary state, and any formatting."
        )


# ---------------------------------------------------------------------------
# Tab 2 — Batch CSV
# ---------------------------------------------------------------------------

def render_batch_tab():
    if not get_template_bytes():
        st.warning("Upload the POA template in the sidebar before generating documents.")
        return

    st.subheader("Batch Generation via CSV")

    with st.expander("Step 1 — Download the CSV template, fill it in Excel", expanded=True):
        st.caption(
            "Download, open in Excel, add one row per case, save as CSV, then upload below. "
            "The first row is a pre-filled example — replace it with your real cases."
        )
        st.download_button(
            "Download CSV Template",
            data=_build_sample_csv(),
            file_name="poa_cases_template.csv",
            mime="text/csv",
        )

    st.markdown("**Step 2 — Upload your completed CSV**")
    uploaded_csv = st.file_uploader("Upload CSV", type=["csv"], label_visibility="collapsed")

    if not uploaded_csv:
        return

    cases, parse_errors = load_cases_from_csv(uploaded_csv.read())

    if parse_errors:
        st.warning(f"{len(parse_errors)} row(s) skipped:")
        for row_num, msg in parse_errors:
            st.caption(f"Row {row_num}: {msg}")

    if not cases:
        st.error("No valid cases found in the CSV.")
        return

    st.success(f"{len(cases)} case(s) found.")
    preview_df = pd.DataFrame(cases)
    show_cols = ["principal_name", "principal_passport", "surrogate_name", "agent_name", "due_date"]
    st.dataframe(
        preview_df[[c for c in show_cols if c in preview_df.columns]],
        use_container_width=True,
    )

    if st.button("Generate All Documents", type="primary", use_container_width=True):
        tpl = get_template_bytes()
        _run_batch_generate(tpl, cases)


def _run_batch_generate(tpl: bytes, cases: list):
    progress = st.progress(0, text="Generating...")
    zip_buf = io.BytesIO()
    failed = []

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, info in enumerate(cases):
            try:
                doc_bytes, _ = generate_poa_bytes(tpl, info)
                zf.writestr(make_filename(info), doc_bytes)
            except Exception as exc:
                failed.append((i + 1, str(exc)))
            progress.progress((i + 1) / len(cases), text=f"{i + 1} / {len(cases)}")

    progress.empty()

    for row, msg in failed:
        st.warning(f"Case {row} failed: {msg}")

    count = len(cases) - len(failed)
    st.success(f"Generated {count} document(s).")
    st.download_button(
        label=f"Download all {count} documents (ZIP)",
        data=zip_buf.getvalue(),
        file_name="POA_Documents.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )


def _build_sample_csv() -> bytes:
    columns = get_csv_columns()
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=columns, lineterminator="\n")
    writer.writeheader()
    writer.writerow(get_example_row())
    return buf.getvalue().encode("utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    st.title("POA Document Generator")
    st.caption("Tsong Law Group, A.P.C.")
    st.divider()

    render_sidebar()

    tab1, tab2 = st.tabs(["Single Case", "Batch (CSV)"])
    with tab1:
        render_single_tab()
    with tab2:
        render_batch_tab()


if __name__ == "__main__":
    main()
