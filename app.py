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
    FIELDS_2IP,
    generate_multi_poa_bytes,
    generate_2ip_poa_bytes,
    validate_case,
    validate_2ip_case,
    make_filename,
    make_filename_2ip,
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
# Template loading — two slots: tpl1 (1 IP) and tpl2 (2 IPs)
# ---------------------------------------------------------------------------

_LOCAL = os.path.dirname(os.path.abspath(__file__))

_LOCAL_TEMPLATES = {
    "tpl1": os.path.join(_LOCAL, "poa_template.docx"),
    "tpl2": os.path.join(_LOCAL, "poa_template_2ip.docx"),
}

_ONEDRIVE_TEMPLATES = {
    "tpl1": (
        r"C:\Users\zhous\OneDrive - Tsong Law Group"
        r"\Ralph Tsong's files - Marketing\POA Pilot"
        r"\Power of Attorney - Shi & Aispuro - C&T Fertility Consultant.docx"
    ),
    "tpl2": (
        r"C:\Users\zhous\OneDrive - Tsong Law Group"
        r"\Ralph Tsong's files - Marketing\POA Pilot"
        r"\poa_template_2ip.docx"
    ),
}


def get_template_bytes(key: str = "tpl1") -> Optional[bytes]:
    if key in st.session_state:
        return st.session_state[key]
    for path in [_LOCAL_TEMPLATES.get(key), _ONEDRIVE_TEMPLATES.get(key)]:
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                data = f.read()
                st.session_state[key] = data
                return data
    return None


# ---------------------------------------------------------------------------
# Sidebar — separate upload per template type
# ---------------------------------------------------------------------------

def render_sidebar():
    with st.sidebar:
        st.header("Templates")

        with st.expander("1 IP + 1 Agent  (Passport-based)", expanded=True):
            tpl1 = get_template_bytes("tpl1")
            if tpl1:
                st.success(f"Ready ({len(tpl1):,} bytes)")
            else:
                st.error("Not loaded")
            up1 = st.file_uploader("Upload template", type=["docx"], key="up_tpl1")
            if up1 and "tpl1" not in st.session_state:
                st.session_state["tpl1"] = up1.read()
                st.session_state.pop("result", None)
            st.caption("`poa_template.docx`  (Shi & Aispuro style)")

        with st.expander("2 IPs + 3 Agents  (CA Driver's License)", expanded=True):
            tpl2 = get_template_bytes("tpl2")
            if tpl2:
                st.success(f"Ready ({len(tpl2):,} bytes)")
            else:
                st.warning("Not loaded — run `python prep_2ip_template.py` first, then upload the output.")
            up2 = st.file_uploader("Upload template", type=["docx"], key="up_tpl2")
            if up2 and "tpl2" not in st.session_state:
                st.session_state["tpl2"] = up2.read()
                st.session_state.pop("result", None)
            st.caption("`poa_template_2ip.docx`  (Ding & Luo style)")


# ---------------------------------------------------------------------------
# Tab 1 — Single case
# ---------------------------------------------------------------------------

def render_single_tab():
    st.subheader("Case Type")

    case_type = st.radio(
        "Select the case structure:",
        options=[
            "1–3 Intended Parents,  1 Agent  (Passport)",
            "2 Intended Parents,  3 Agents  (CA Driver's License)",
        ],
        key="case_type_selector",
        horizontal=True,
    )

    is_2ip = "3 Agents" in case_type

    # Clear stale result when the user switches case types
    if st.session_state.get("_last_case_type") != case_type:
        st.session_state["_last_case_type"] = case_type
        st.session_state.pop("result", None)

    tpl_key = "tpl2" if is_2ip else "tpl1"
    if not get_template_bytes(tpl_key):
        tpl_label = "2-IP + 3-Agent" if is_2ip else "1-IP + 1-Agent"
        st.warning(f"Upload the {tpl_label} template in the sidebar first.")
        return

    st.divider()

    if is_2ip:
        _render_2ip_form()
    else:
        _render_1ip_form()


# ---- 1-IP form (existing logic) -------------------------------------------

def _render_1ip_form():
    st.subheader("Case Information")

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

        for p_idx in range(num_principals):
            label_num = f"Principal {p_idx + 1}" if num_principals > 1 else "Principal (Intended Parent)"
            st.markdown(f"#### {label_num}")
            p_left, p_right = st.columns(2)

            with p_left:
                for field in PRINCIPAL_FIELDS[:2]:
                    lbl = field["label"] + (" *" if field["required"] else "")
                    st.text_input(lbl, value=field["default"],
                                  key=f"p{p_idx}_{field['key']}", help=field["help"] or None)

            with p_right:
                for field in PRINCIPAL_FIELDS[2:]:
                    lbl = field["label"] + (" *" if field["required"] else "")
                    st.text_input(lbl, value=field["default"],
                                  key=f"p{p_idx}_{field['key']}", help=field["help"] or None)

            st.file_uploader(
                f"{'Principal' if num_principals == 1 else label_num} — Passport / ID Photo",
                type=["jpg", "jpeg", "png"],
                key=f"photo_p{p_idx}",
            )
            if p_idx < num_principals - 1:
                st.divider()

        st.divider()
        st.markdown("**Agent's Passport / ID Photo**")
        st.file_uploader("Agent / Attorney-in-Fact ID photo", type=["jpg", "jpeg", "png"],
                         key="photo_agent", label_visibility="collapsed")

        st.write("")
        submitted = st.form_submit_button("Generate POA Document", type="primary",
                                          use_container_width=True)

    if submitted:
        _run_1ip_generate(num_principals)

    _render_result()


def _run_1ip_generate(num_principals: int):
    tpl = get_template_bytes("tpl1")
    if not tpl:
        st.session_state["result"] = {"error": "No template loaded."}
        return

    case_info = {}
    for field in CASE_FIELDS:
        raw = st.session_state.get(f"case_{field['key']}", field["default"]) or ""
        case_info[field["key"]] = raw.strip().upper() if field["to_upper"] else raw.strip()

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

    agent_file = st.session_state.get("photo_agent")
    agent_photo = agent_file.read() if agent_file is not None else None

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


# ---- 2-IP form (new) -------------------------------------------------------

def _render_2ip_form():
    st.subheader("Case Information  —  2 IPs / 3 Agents")

    with st.form("poa_2ip_form", clear_on_submit=False):

        # IPs
        for ip_num in (1, 2):
            st.markdown(f"#### Intended Parent {ip_num}")
            c1, c2 = st.columns(2)
            with c1:
                st.text_input(f"Full Name *", key=f"ip{ip_num}_name",
                              help="ALL CAPS — e.g., CHENGFANG DING")
            with c2:
                st.text_input(f"Passport Number *", key=f"ip{ip_num}_passport",
                              help="e.g., ED3297013")
            st.file_uploader(f"IP{ip_num} Passport / ID Photo",
                             type=["jpg", "jpeg", "png"], key=f"photo_ip{ip_num}")
            st.divider()

        # Case details
        st.markdown("#### Case Details")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Child Last Name *", key="2ip_child_last_name",
                          help="ALL CAPS — will appear as 'Infant [LAST NAME]'")
            st.text_input("Surrogate Full Name *", key="2ip_surrogate_name", help="ALL CAPS")
            st.text_input("Surrogate Date of Birth *", key="2ip_surrogate_dob",
                          help="e.g., May 22, 1996")
        with c2:
            st.text_input("Due Date *", key="2ip_due_date", help="e.g., March 15, 2026")
            st.text_input("Hospital Name", key="2ip_hospital_name",
                          value="Loma Linda University Medical Center")

        st.divider()

        # Agents
        for a in range(1, 4):
            st.markdown(f"#### Agent {a}")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Full Name *", key=f"2ip_agent{a}_name", help="ALL CAPS")
            with c2:
                st.text_input("Date of Birth *", key=f"2ip_agent{a}_dob",
                              help="e.g., 02/24/1986")
            with c3:
                st.text_input("CA Driver License No. *", key=f"2ip_agent{a}_dl",
                              help="e.g., Y1234567")
            st.file_uploader(f"Agent {a} ID Photo",
                             type=["jpg", "jpeg", "png"], key=f"photo_agent{a}")
            if a < 3:
                st.divider()

        st.divider()

        # Firm
        st.markdown("#### Firm Details")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Handling Attorney", key="2ip_attorney_name", value="Xuelan Fang")
        with c2:
            st.text_input("Fertility Agency (for filename)", key="2ip_agency_name")

        st.write("")
        submitted = st.form_submit_button("Generate 2-IP POA Document", type="primary",
                                          use_container_width=True)

    if submitted:
        _run_2ip_generate()

    _render_result()


def _run_2ip_generate():
    tpl = get_template_bytes("tpl2")
    if not tpl:
        st.session_state["result"] = {"error": "No 2-IP template loaded."}
        return

    def _get(key, upper=False, default=""):
        raw = (st.session_state.get(key) or default).strip()
        return raw.upper() if upper else raw

    info = {
        "ip1_name":        _get("ip1_name", upper=True),
        "ip1_passport":    _get("ip1_passport"),
        "ip2_name":        _get("ip2_name", upper=True),
        "ip2_passport":    _get("ip2_passport"),
        "child_last_name": _get("2ip_child_last_name", upper=True),
        "surrogate_name":  _get("2ip_surrogate_name", upper=True),
        "surrogate_dob":   _get("2ip_surrogate_dob"),
        "due_date":        _get("2ip_due_date"),
        "hospital_name":   _get("2ip_hospital_name", default="Loma Linda University Medical Center"),
        "agent1_name":     _get("2ip_agent1_name", upper=True),
        "agent1_dob":      _get("2ip_agent1_dob"),
        "agent1_dl":       _get("2ip_agent1_dl"),
        "agent2_name":     _get("2ip_agent2_name", upper=True),
        "agent2_dob":      _get("2ip_agent2_dob"),
        "agent2_dl":       _get("2ip_agent2_dl"),
        "agent3_name":     _get("2ip_agent3_name", upper=True),
        "agent3_dob":      _get("2ip_agent3_dob"),
        "agent3_dl":       _get("2ip_agent3_dl"),
        "attorney_name":   _get("2ip_attorney_name", default="Xuelan Fang"),
        "agency_name":     _get("2ip_agency_name"),
    }

    errors = validate_2ip_case(info)
    if errors:
        st.session_state["result"] = {"errors": errors}
        return

    photos = {}
    for agent_num in (1, 2, 3):
        f = st.session_state.get(f"photo_agent{agent_num}")
        if f is not None:
            photos[f"agent{agent_num}"] = f.read()
    for ip_num in (1, 2):
        f = st.session_state.get(f"photo_ip{ip_num}")
        if f is not None:
            photos[f"ip{ip_num}"] = f.read()

    try:
        doc_bytes, count = generate_2ip_poa_bytes(tpl, info, photos=photos)
        filename = make_filename_2ip(info)
        st.session_state["result"] = {"doc_bytes": doc_bytes, "filename": filename, "count": count}
    except Exception as exc:
        st.session_state["result"] = {"error": str(exc)}


# ---- Shared result renderer ------------------------------------------------

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
        st.caption("Review after downloading: spelling, dates, and any formatting.")


# ---------------------------------------------------------------------------
# Tab 2 — Batch CSV  (1-IP cases only)
# ---------------------------------------------------------------------------

def render_batch_tab():
    if not get_template_bytes("tpl1"):
        st.warning("Upload the 1-IP + 1-Agent template in the sidebar first.")
        return

    st.subheader("Batch Generation via CSV")
    st.caption("For single-principal cases only. Use the Single Case tab for 2-IP cases.")

    with st.expander("Step 1 — Download the CSV template, fill it in Excel", expanded=True):
        st.caption("One row per case. The first row is a pre-filled example.")
        st.download_button("Download CSV Template", data=_build_sample_csv(),
                           file_name="poa_cases_template.csv", mime="text/csv")

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
        _run_batch(get_template_bytes("tpl1"), cases)


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
