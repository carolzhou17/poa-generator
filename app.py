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
    FIELDS, PRINCIPAL_FIELDS, CASE_FIELDS,
    generate_multi_poa_bytes,
    validate_case,
    make_filename,
    load_cases_from_csv,
    get_csv_columns,
    get_example_row,
)

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


def get_template_bytes() -> Optional[bytes]:
    if "template_bytes" in st.session_state:
        return st.session_state["template_bytes"]
    for path in [LOCAL_TEMPLATE, ONEDRIVE_TEMPLATE]:
        if os.path.exists(path):
            with open(path, "rb") as f:
                data = f.read()
                st.session_state["template_bytes"] = data
                return data
    return None


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

def render_sidebar():
    with st.sidebar:
        st.header("Template")
        tpl = get_template_bytes()
        if tpl:
            st.success(f"Template ready ({len(tpl):,} bytes)")
        else:
            st.error("No template loaded — upload one below.")

        uploaded = st.file_uploader("Upload POA template (.docx)", type=["docx"])
        if uploaded and "template_bytes" not in st.session_state:
            st.session_state["template_bytes"] = uploaded.read()
            st.session_state.pop("result", None)

        st.divider()
        st.caption(
            "Upload:  \n"
            "`Power of Attorney - Shi & Aispuro...docx`  \n"
            "from the POA Pilot folder on OneDrive."
        )


# ---------------------------------------------------------------------------
# Tab 1 — Single case (supports 1–3 principals)
# ---------------------------------------------------------------------------

def render_single_tab():
    if not get_template_bytes():
        st.warning("Upload the POA template in the sidebar first.")
        return

    st.subheader("Case Information")

    # Number of principals selector — outside the form so it re-renders sections
    num_principals = st.selectbox(
        "How many Intended Parents (principals) in this case?",
        options=[1, 2, 3],
        key="num_principals",
        help="Each principal gets their own POA section in the combined document.",
    )

    sections = {}
    for f in CASE_FIELDS:
        sections.setdefault(f["section"], []).append(f)

    with st.form("poa_form", clear_on_submit=False):

        # ---- Shared case fields ----
        st.markdown("#### Case Details")
        left, right = st.columns(2)
        section_list = list(sections.items())
        for i, (sec_name, fields) in enumerate(section_list):
            col = left if i % 2 == 0 else right
            with col:
                st.markdown(f"**{sec_name}**")
                for field in fields:
                    label = field["label"] + (" *" if field["required"] else "")
                    st.text_input(
                        label,
                        value=field["default"],
                        key=f"case_{field['key']}",
                        help=field["help"] or None,
                    )
                st.write("")

        st.divider()

        # ---- Per-principal sections ----
        for p_idx in range(num_principals):
            label_num = f"Principal {p_idx + 1}" if num_principals > 1 else "Principal (Intended Parent)"
            st.markdown(f"#### {label_num}")
            p_left, p_right = st.columns(2)

            with p_left:
                for field in PRINCIPAL_FIELDS[:2]:   # name, role
                    lbl = field["label"] + (" *" if field["required"] else "")
                    st.text_input(
                        lbl,
                        value=field["default"],
                        key=f"p{p_idx}_{field['key']}",
                        help=field["help"] or None,
                    )

            with p_right:
                for field in PRINCIPAL_FIELDS[2:]:   # passport country, passport number
                    lbl = field["label"] + (" *" if field["required"] else "")
                    st.text_input(
                        lbl,
                        value=field["default"],
                        key=f"p{p_idx}_{field['key']}",
                        help=field["help"] or None,
                    )

            # ID photo for this principal
            st.file_uploader(
                f"{'Principal' if num_principals == 1 else label_num} — Passport / ID Photo",
                type=["jpg", "jpeg", "png"],
                key=f"photo_p{p_idx}",
            )
            if p_idx < num_principals - 1:
                st.divider()

        st.divider()

        # ---- Agent photo (shared across all principals) ----
        st.markdown("**Agent's Passport / ID Photo**")
        st.file_uploader(
            "Agent / Attorney-in-Fact ID photo",
            type=["jpg", "jpeg", "png"],
            key="photo_agent",
            label_visibility="collapsed",
        )

        st.write("")
        submitted = st.form_submit_button(
            "Generate POA Document",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        _run_generate(num_principals)

    _render_result()


def _run_generate(num_principals: int):
    tpl = get_template_bytes()
    if not tpl:
        st.session_state["result"] = {"error": "No template loaded."}
        return

    # Read shared case fields
    case_info = {}
    for field in CASE_FIELDS:
        raw = st.session_state.get(f"case_{field['key']}", field["default"]) or ""
        case_info[field["key"]] = raw.strip().upper() if field["to_upper"] else raw.strip()

    # Read per-principal data
    principals = []
    for p_idx in range(num_principals):
        p = {}
        for field in PRINCIPAL_FIELDS:
            raw = st.session_state.get(f"p{p_idx}_{field['key']}", field["default"]) or ""
            p[field["key"]] = raw.strip().upper() if field["to_upper"] else raw.strip()

        photo_file = st.session_state.get(f"photo_p{p_idx}")
        if photo_file is not None:
            p["photo"] = photo_file.read()

        principals.append(p)

    # Read agent photo
    agent_file = st.session_state.get("photo_agent")
    agent_photo = agent_file.read() if agent_file is not None else None

    # Validate
    errors = validate_case(case_info, principals)
    if errors:
        st.session_state["result"] = {"errors": errors}
        return

    try:
        doc_bytes, count = generate_multi_poa_bytes(tpl, case_info, principals, agent_photo)
        filename = make_filename(case_info, principals)
        st.session_state["result"] = {"doc_bytes": doc_bytes, "filename": filename, "count": count}
    except Exception as exc:
        st.session_state["result"] = {"error": str(exc)}


def _render_result():
    result = st.session_state.get("result")
    if not result:
        return
    if "error" in result:
        st.error(result["error"])
    elif "errors" in result:
        for e in result["errors"]:
            st.error(e)
    else:
        n = result["count"]
        st.success(f"Document ready — {n} field replacement{'s' if n != 1 else ''} made.")
        st.download_button(
            label=f"Download: {result['filename']}",
            data=result["doc_bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
        st.caption(
            "Review after downloading: agent pronouns (her/his), "
            "notary state, and any formatting."
        )


# ---------------------------------------------------------------------------
# Tab 2 — Batch CSV
# ---------------------------------------------------------------------------

def render_batch_tab():
    if not get_template_bytes():
        st.warning("Upload the POA template in the sidebar first.")
        return

    st.subheader("Batch Generation via CSV")
    st.caption("For single-principal cases only. Use the Single Case tab for multi-principal cases.")

    with st.expander("Step 1 — Download the CSV template, fill it in Excel", expanded=True):
        st.caption(
            "One row per case. The first row is a pre-filled example — replace it with your data."
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
    show = ["principal_name", "principal_passport", "surrogate_name", "agent_name", "due_date"]
    st.dataframe(preview_df[[c for c in show if c in preview_df.columns]], use_container_width=True)

    if st.button("Generate All Documents", type="primary", use_container_width=True):
        _run_batch(get_template_bytes(), cases)


def _run_batch(tpl: bytes, cases: list):
    from poa_generator import generate_poa_bytes
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
            progress.progress((i + 1) / len(cases), text=f"{i+1} / {len(cases)}")
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
