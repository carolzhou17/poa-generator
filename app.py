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
    # Legacy (batch CSV only)
    FIELDS, CASE_FIELDS,
    generate_poa_bytes,
    validate_case,
    make_filename,
    load_cases_from_csv,
    get_csv_columns,
    get_example_row,
    # V2 unified
    generate_poa_v2_bytes,
    validate_v2,
    make_filename_v2,
)

st.set_page_config(
    page_title="POA Generator — Tsong Law Group",
    page_icon="⚖️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Agent ID type choices  (label shown in dropdown → full phrase for document)
# ---------------------------------------------------------------------------

ID_TYPES = {
    "California Driver's License": "California’s Driver License No.",
    "Nevada Driver's License":     "Nevada’s Driver License No.",
    "Texas Driver's License":      "Texas Driver’s License No.",
    "New York Driver's License":   "New York Driver’s License No.",
    "Illinois Driver's License":   "Illinois Driver’s License No.",
    "Washington Driver's License": "Washington Driver’s License No.",
    "Oregon Driver's License":     "Oregon Driver’s License No.",
    "China Passport":              "People’s Republic of China Passport No.",
    "Taiwan Passport":             "Republic of China Passport No.",
    "US Passport":                 "United States Passport No.",
    "Other (type full phrase)":    None,
}

_ID_LABELS = list(ID_TYPES.keys())

# ---------------------------------------------------------------------------
# Template loading — one slot per combo (num_ips, num_agents)
# ---------------------------------------------------------------------------

_LOCAL = os.path.dirname(os.path.abspath(__file__))


def _combo_key(num_ips: int, num_agents: int) -> str:
    return f"tpl_{num_ips}_{num_agents}"


def _local_path(num_ips: int, num_agents: int) -> str:
    return os.path.join(_LOCAL, f"poa_template_{num_ips}ip_{num_agents}agent.docx")


def get_template_bytes(num_ips: int, num_agents: int) -> Optional[bytes]:
    key = _combo_key(num_ips, num_agents)
    if key in st.session_state:
        return st.session_state[key]
    path = _local_path(num_ips, num_agents)
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        st.session_state[key] = data
        return data
    return None


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

_COMBO_LABELS = {
    (1, 1): "1 IP  /  1 Agent",
    (1, 2): "1 IP  /  2 Agents",
    (1, 3): "1 IP  /  3 Agents",
    (2, 1): "2 IPs  /  1 Agent",
    (2, 2): "2 IPs  /  2 Agents",
    (2, 3): "2 IPs  /  3 Agents",
}


def render_sidebar():
    with st.sidebar:
        st.header("Templates")
        ready = sum(1 for (i, a) in _COMBO_LABELS if get_template_bytes(i, a))
        total = len(_COMBO_LABELS)
        if ready == total:
            st.success(f"All {total} templates loaded")
        else:
            st.warning(f"{ready} / {total} templates loaded")
        for (i, a), lbl in _COMBO_LABELS.items():
            icon = "✓" if get_template_bytes(i, a) else "✗"
            st.caption(f"{icon}  {lbl}")


# ---------------------------------------------------------------------------
# Tab 1 — Single case (V2, all 6 combos)
# ---------------------------------------------------------------------------

def render_single_tab():
    c1, c2 = st.columns(2)
    with c1:
        num_ips = st.selectbox("Intended Parents:", [1, 2], key="v2_num_ips")
    with c2:
        num_agents = st.selectbox("Agents / POAs:", [1, 2, 3], key="v2_num_agents")

    # Clear stale result when combo changes
    combo_key = f"{num_ips}_{num_agents}"
    if st.session_state.get("_v2_last_combo") != combo_key:
        st.session_state["_v2_last_combo"] = combo_key
        st.session_state.pop("result", None)

    tpl = get_template_bytes(num_ips, num_agents)
    label = _COMBO_LABELS[(num_ips, num_agents)]
    if not tpl:
        st.error(f"Template for **{label}** not found. Run `python prep_all_templates.py` in the app folder.")
        return

    st.divider()

    with st.form("poa_v2_form", clear_on_submit=False):

        # ── IP sections ──────────────────────────────────────────────────────
        for n in range(1, num_ips + 1):
            header = f"#### Intended Parent {n}" if num_ips > 1 else "#### Intended Parent"
            st.markdown(header)
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input(
                    "Full Name *",
                    key=f"v2_ip{n}_name",
                    help="ALL CAPS — e.g., CHENGFANG DING",
                )
            with c2:
                st.selectbox(
                    "Role",
                    ["Intended Father", "Intended Mother"],
                    key=f"v2_ip{n}_role",
                )
            with c3:
                st.text_input(
                    "Passport Country",
                    value="People's Republic of China",
                    key=f"v2_ip{n}_country",
                )
            c4, c5 = st.columns(2)
            with c4:
                st.text_input(
                    "Passport Number *",
                    key=f"v2_ip{n}_passport",
                    help="e.g., ED3297013",
                )
            with c5:
                st.file_uploader(
                    f"IP {n} — Passport / ID Photo",
                    type=["jpg", "jpeg", "png"],
                    key=f"v2_photo_ip{n}",
                )
            st.divider()

        # ── Case details ─────────────────────────────────────────────────────
        st.markdown("#### Case Details")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("Child Last Name *", key="v2_child_last_name",
                          help="ALL CAPS — appears as 'Infant [NAME]'")
            st.text_input("Due Date *", key="v2_due_date",
                          help="e.g., March 15, 2026")
        with c2:
            st.text_input("Surrogate Full Name *", key="v2_surrogate_name",
                          help="ALL CAPS")
            st.text_input("Hospital Name", key="v2_hospital_name",
                          value="Loma Linda University Medical Center")
        with c3:
            st.text_input("Surrogate Date of Birth *", key="v2_surrogate_dob",
                          help="e.g., May 22, 1996")
            c3a, c3b = st.columns(2)
            with c3a:
                st.text_input("Hospital City", key="v2_hospital_city",
                              value="Loma Linda")
            with c3b:
                st.text_input("Hospital State", key="v2_hospital_state",
                              value="California")

        st.divider()

        # ── Agent sections ────────────────────────────────────────────────────
        for n in range(1, num_agents + 1):
            header = f"#### Agent {n}" if num_agents > 1 else "#### Agent / POA"
            st.markdown(header)
            c1, c2 = st.columns(2)
            with c1:
                st.text_input("Full Name *", key=f"v2_agent{n}_name", help="ALL CAPS")
                st.text_input("Date of Birth *", key=f"v2_agent{n}_dob",
                              help="e.g., 02/24/1986")
            with c2:
                id_type_choice = st.selectbox(
                    "ID Type",
                    _ID_LABELS,
                    key=f"v2_agent{n}_id_type",
                )
                id_number = st.text_input("ID Number *", key=f"v2_agent{n}_id_number",
                                          help="e.g., Y1234567 or EJ4979964")
                if ID_TYPES[id_type_choice] is None:
                    st.text_input(
                        "Full ID type label *",
                        key=f"v2_agent{n}_id_custom",
                        help='e.g., "Florida\'s Driver License No."',
                    )
            st.file_uploader(
                f"Agent {n} — ID Photo",
                type=["jpg", "jpeg", "png"],
                key=f"v2_photo_agent{n}",
            )
            if n < num_agents:
                st.divider()

        st.divider()

        # ── Firm ─────────────────────────────────────────────────────────────
        st.markdown("#### Firm")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Handling Attorney", key="v2_attorney_name",
                          value="Xuelan Fang")
        with c2:
            st.text_input("Fertility Agency (for filename only)",
                          key="v2_agency_name")

        st.write("")
        submitted = st.form_submit_button(
            "Generate POA Document", type="primary", use_container_width=True
        )

    if submitted:
        _run_v2_generate(num_ips, num_agents)

    _render_result()


def _run_v2_generate(num_ips: int, num_agents: int):
    tpl = get_template_bytes(num_ips, num_agents)
    if not tpl:
        st.session_state["result"] = {"error": "No template loaded."}
        return

    def _get(k, upper=False, default=""):
        raw = (st.session_state.get(k) or default).strip()
        return raw.upper() if upper else raw

    info: dict = {}

    # IPs
    for n in range(1, num_ips + 1):
        info[f"ip{n}_name"]    = _get(f"v2_ip{n}_name", upper=True)
        info[f"ip{n}_role"]    = _get(f"v2_ip{n}_role", default="Intended Father")
        info[f"ip{n}_country"] = _get(f"v2_ip{n}_country",
                                      default="People's Republic of China")
        info[f"ip{n}_passport"] = _get(f"v2_ip{n}_passport")

    # Case
    info["child_last_name"] = _get("v2_child_last_name", upper=True)
    info["surrogate_name"]  = _get("v2_surrogate_name",  upper=True)
    info["surrogate_dob"]   = _get("v2_surrogate_dob")
    info["due_date"]        = _get("v2_due_date")
    info["hospital_name"]   = _get("v2_hospital_name",
                                   default="Loma Linda University Medical Center")
    info["hospital_city"]   = _get("v2_hospital_city",  default="Loma Linda")
    info["hospital_state"]  = _get("v2_hospital_state", default="California")
    info["attorney_name"]   = _get("v2_attorney_name",  default="Xuelan Fang")
    info["agency_name"]     = _get("v2_agency_name")

    # Agents — build the full ID phrase here
    for n in range(1, num_agents + 1):
        info[f"agent{n}_name"] = _get(f"v2_agent{n}_name", upper=True)
        info[f"agent{n}_dob"]  = _get(f"v2_agent{n}_dob")
        id_choice = _get(f"v2_agent{n}_id_type",
                         default=_ID_LABELS[0])
        phrase = ID_TYPES.get(id_choice)
        if phrase is None:
            phrase = _get(f"v2_agent{n}_id_custom")
        id_num = _get(f"v2_agent{n}_id_number")
        info[f"agent{n}_id"] = f"{phrase} {id_num}".strip() if phrase else id_num

    errors = validate_v2(num_ips, num_agents, info)
    if errors:
        st.session_state["result"] = {"errors": errors}
        return

    photos: dict = {}
    for n in range(1, num_ips + 1):
        f = st.session_state.get(f"v2_photo_ip{n}")
        if f is not None:
            photos[f"ip{n}"] = f.read()
    for n in range(1, num_agents + 1):
        f = st.session_state.get(f"v2_photo_agent{n}")
        if f is not None:
            photos[f"agent{n}"] = f.read()

    try:
        doc_bytes, count = generate_poa_v2_bytes(
            tpl, num_ips, num_agents, info, photos=photos or None
        )
        filename = make_filename_v2(num_ips, info)
        st.session_state["result"] = {
            "doc_bytes": doc_bytes, "filename": filename, "count": count
        }
    except Exception as exc:
        st.session_state["result"] = {"error": str(exc)}


# ---------------------------------------------------------------------------
# Shared result renderer
# ---------------------------------------------------------------------------

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
            label=f"⬇  {result['filename']}",
            data=result["doc_bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
        st.caption("Review the document after downloading: dates, spelling, and formatting.")


# ---------------------------------------------------------------------------
# Tab 2 — Batch CSV  (legacy 1-IP / 1-agent)
# ---------------------------------------------------------------------------

def render_batch_tab():
    tpl = get_template_bytes(1, 1)
    if not tpl:
        st.error("Template for 1 IP / 1 Agent not found. Run `python prep_all_templates.py`.")
        return

    st.subheader("Batch Generation via CSV")
    st.caption("Single-principal cases only. Use the Single Case tab for other combos.")

    with st.expander("Step 1 — Download the CSV template", expanded=True):
        st.download_button(
            "Download CSV Template",
            data=_build_sample_csv(),
            file_name="poa_cases_template.csv",
            mime="text/csv",
        )

    st.markdown("**Step 2 — Upload your completed CSV**")
    uploaded_csv = st.file_uploader("Upload CSV", type=["csv"],
                                    label_visibility="collapsed")
    if not uploaded_csv:
        return

    cases, parse_errors = load_cases_from_csv(uploaded_csv.read())
    if parse_errors:
        st.warning(f"{len(parse_errors)} row(s) skipped:")
        for row_num, msg in parse_errors:
            st.caption(f"Row {row_num}: {msg}")

    if not cases:
        st.error("No valid cases found.")
        return

    st.success(f"{len(cases)} case(s) found.")
    preview_df = pd.DataFrame(cases)
    show = ["principal_name", "principal_passport", "surrogate_name", "agent_name", "due_date"]
    st.dataframe(preview_df[[c for c in show if c in preview_df.columns]],
                 use_container_width=True)

    if st.button("Generate All Documents", type="primary", use_container_width=True):
        _run_batch(tpl, cases)


def _run_batch(tpl: bytes, cases: list):
    progress = st.progress(0, text="Generating…")
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
