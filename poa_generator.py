"""
poa_generator.py  —  Pure logic, no UI.
Handles field definitions, Word replacement, CSV parsing, and document generation.
"""

import io
import csv
from typing import Optional, List
from docx import Document
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Field definitions
# ---------------------------------------------------------------------------
# per_principal=True  → belongs to one principal (repeated for each IP)
# per_principal=False → shared across the whole case (surrogate, agent, etc.)

FIELDS = [
    # ---- Per-principal fields ----
    {
        "key":           "principal_name",
        "label":         "Full Name",
        "section":       "Principal",
        "template_value":"CHENGFANG SHI",
        "required":      True,
        "to_upper":      True,
        "default":       "",
        "help":          "ALL CAPS — e.g., JOHN SMITH",
        "per_principal": True,
    },
    {
        "key":           "principal_role",
        "label":         "Role",
        "section":       "Principal",
        "template_value":"Intended Father",
        "required":      False,
        "to_upper":      False,
        "default":       "Intended Father",
        "help":          "Intended Father  or  Intended Mother",
        "per_principal": True,
    },
    {
        "key":           "passport_country",
        "label":         "Passport Country",
        "section":       "Principal",
        "template_value":"People’s Republic of China",
        "required":      False,
        "to_upper":      False,
        "default":       "People’s Republic of China",
        "help":          "Country as it appears on passport",
        "per_principal": True,
    },
    {
        "key":           "principal_passport",
        "label":         "Passport Number",
        "section":       "Principal",
        "template_value":"ED3297013",
        "required":      True,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., ED3297013",
        "per_principal": True,
    },

    # ---- Shared case fields ----
    {
        "key":           "child_last_name",
        "label":         "Child Last Name",
        "section":       "Child",
        "template_value":"SHI",
        "required":      True,
        "to_upper":      True,
        "default":       "",
        "help":          "Will appear as 'infant CHILD [LAST NAME]'",
        "per_principal": False,
    },
    {
        "key":           "surrogate_name",
        "label":         "Surrogate Full Name",
        "section":       "Surrogate",
        "template_value":"SELENA MARIA AISPURO",
        "required":      True,
        "to_upper":      True,
        "default":       "",
        "help":          "ALL CAPS",
        "per_principal": False,
    },
    {
        "key":           "surrogate_dob",
        "label":         "Surrogate Date of Birth",
        "section":       "Surrogate",
        "template_value":"May 22, 1996",
        "required":      True,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., May 22, 1996",
        "per_principal": False,
    },
    {
        "key":           "due_date",
        "label":         "Due Date",
        "section":       "Birth Details",
        "template_value":"December 4, 2025",
        "required":      True,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., March 15, 2026",
        "per_principal": False,
    },
    {
        "key":           "hospital_name",
        "label":         "Hospital Name",
        "section":       "Birth Details",
        "template_value":"Loma Linda University Medical Center",
        "required":      False,
        "to_upper":      False,
        "default":       "Loma Linda University Medical Center",
        "help":          "Full hospital name",
        "per_principal": False,
    },
    {
        "key":           "hospital_city",
        "label":         "Hospital City",
        "section":       "Birth Details",
        "template_value":"Loma Linda",
        "required":      False,
        "to_upper":      False,
        "default":       "Loma Linda",
        "help":          "",
        "per_principal": False,
    },
    {
        "key":           "hospital_state",
        "label":         "Hospital State",
        "section":       "Birth Details",
        "template_value":"California",
        "required":      False,
        "to_upper":      False,
        "default":       "California",
        "help":          "",
        "per_principal": False,
    },
    {
        "key":           "agent_name",
        "label":         "Agent / Attorney-in-Fact Full Name",
        "section":       "Agent",
        "template_value":"JIAJIA GAO",
        "required":      True,
        "to_upper":      True,
        "default":       "",
        "help":          "ALL CAPS",
        "per_principal": False,
    },
    {
        "key":           "agent_dob",
        "label":         "Agent Date of Birth",
        "section":       "Agent",
        "template_value":"02/24/1986",
        "required":      True,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., 02/24/1986",
        "per_principal": False,
    },
    {
        "key":           "agent_passport",
        "label":         "Agent Passport Number",
        "section":       "Agent",
        "template_value":"EJ4979964",
        "required":      True,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., EJ4979964",
        "per_principal": False,
    },
    {
        "key":           "agent_pronoun",
        "label":         "Agent Pronoun",
        "section":       "Agent",
        "template_value":"her",
        "required":      False,
        "to_upper":      False,
        "default":       "her",
        "help":          "her / his / their",
        "per_principal": False,
    },
    {
        "key":           "attorney_name",
        "label":         "Handling Attorney",
        "section":       "Firm",
        "template_value":"Xuelan Fang",
        "required":      False,
        "to_upper":      False,
        "default":       "Xuelan Fang",
        "help":          "",
        "per_principal": False,
    },
    {
        "key":           "agency_name",
        "label":         "Fertility Agency (for filename)",
        "section":       "Firm",
        "template_value":"",
        "required":      False,
        "to_upper":      False,
        "default":       "",
        "help":          "e.g., C&T Fertility Consultant",
        "per_principal": False,
    },
]

PRINCIPAL_FIELDS = [f for f in FIELDS if f["per_principal"]]
CASE_FIELDS      = [f for f in FIELDS if not f["per_principal"]]

# Replacement order: specific strings before the substrings they contain.
REPLACEMENT_ORDER = [
    "hospital_name",       # "Loma Linda University Medical Center" before "Loma Linda"
    "hospital_city",
    "hospital_state",
    "due_date",
    "surrogate_dob",
    "agent_dob",
    "principal_name",
    "child_last_name",     # special: replaces "CHILD SHI", not just "SHI"
    "surrogate_name",
    "agent_name",
    "attorney_name",
    "principal_passport",
    "agent_passport",
    "passport_country",
    "principal_role",
    "agent_pronoun",
]


# ---------------------------------------------------------------------------
# Low-level Word replacement engine
# ---------------------------------------------------------------------------

def _replace_once(para, old: str, new: str) -> bool:
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return False
    idx, end_idx = full.index(old), full.index(old) + len(old)
    pos = 0
    s_run = s_off = e_run = e_off = None
    for i, run in enumerate(para.runs):
        run_end = pos + len(run.text)
        if s_run is None and run_end > idx:
            s_run, s_off = i, idx - pos
        if e_run is None and run_end >= end_idx:
            e_run, e_off = i, end_idx - pos
            break
        pos += len(run.text)
    if s_run is None or e_run is None:
        return False
    runs = list(para.runs)
    if s_run == e_run:
        r = runs[s_run]
        r.text = r.text[:s_off] + new + r.text[e_off:]
    else:
        runs[s_run].text = runs[s_run].text[:s_off] + new
        for i in range(s_run + 1, e_run):
            runs[i].text = ""
        runs[e_run].text = runs[e_run].text[e_off:]
    return True


def _replace_in_paragraph(para, old: str, new: str) -> int:
    count = 0
    while _replace_once(para, old, new):
        count += 1
    return count


def _replace_in_doc(doc, old: str, new: str) -> int:
    if not old or old == new:
        return 0
    count = 0
    for para in doc.paragraphs:
        count += _replace_in_paragraph(para, old, new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    count += _replace_in_paragraph(para, old, new)
    return count


# ---------------------------------------------------------------------------
# Image replacement
# ---------------------------------------------------------------------------

def _find_image_paragraphs(doc) -> list:
    return [p for p in doc.paragraphs
            if p._element.find('.//' + qn('a:blip')) is not None]


def _replace_image(doc, para, new_bytes: bytes) -> bool:
    blip = para._element.find('.//' + qn('a:blip'))
    if blip is None:
        return False
    rid = blip.get(qn('r:embed'))
    if rid is None:
        return False
    doc.part.related_parts[rid]._blob = new_bytes
    return True


# ---------------------------------------------------------------------------
# Document merging
# ---------------------------------------------------------------------------

def _merge_documents(doc_bytes_list: List[bytes]) -> bytes:
    """Combine multiple .docx files into one, each starting on a new page."""
    if len(doc_bytes_list) == 1:
        return doc_bytes_list[0]

    from docxcompose.composer import Composer

    base = Document(io.BytesIO(doc_bytes_list[0]))
    composer = Composer(base)

    for extra_bytes in doc_bytes_list[1:]:
        src = Document(io.BytesIO(extra_bytes))
        composer.append(src)

    out = io.BytesIO()
    composer.save(out)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def validate_case(case_info: dict, principals: Optional[List[dict]] = None) -> list:
    """Return list of error strings. Empty = valid."""
    errors = []
    # Validate shared case fields
    for field in CASE_FIELDS:
        if field["required"] and not case_info.get(field["key"], "").strip():
            errors.append(f"'{field['label']}' is required.")
    # Validate each principal
    if principals:
        for i, p in enumerate(principals, 1):
            for field in PRINCIPAL_FIELDS:
                if field["required"] and not p.get(field["key"], "").strip():
                    errors.append(f"Principal {i}: '{field['label']}' is required.")
    return errors


def make_filename(case_info: dict, principals: Optional[List[dict]] = None) -> str:
    if principals:
        last_names = [p.get("principal_name", "Client").split()[-1] for p in principals]
    else:
        last_names = [case_info.get("principal_name", "Client").split()[-1]]

    surrogate_last = case_info.get("surrogate_name", "Surrogate").split()[-1]
    agency = case_info.get("agency_name", "").strip()

    principals_str = " & ".join(last_names)
    parts = [f"POA - {principals_str} & {surrogate_last}"]
    if agency:
        parts.append(agency)
    return " - ".join(parts) + ".docx"


def generate_poa_bytes(
    template_bytes: bytes,
    info: dict,
    principal_photo: Optional[bytes] = None,
    agent_photo: Optional[bytes] = None,
) -> tuple:
    """Generate one POA document (single principal). Returns (bytes, count)."""
    doc = Document(io.BytesIO(template_bytes))
    total = 0

    for key in REPLACEMENT_ORDER:
        field = next((f for f in FIELDS if f["key"] == key), None)
        if field is None:
            continue
        value = info.get(key, "").strip()
        if field["to_upper"]:
            value = value.upper()
        if not value:
            value = field["default"]
        template_val = field["template_value"]

        if key == "child_last_name":
            total += _replace_in_doc(doc, f"CHILD {template_val}", f"CHILD {value}")
        else:
            total += _replace_in_doc(doc, template_val, value)

    img_paras = _find_image_paragraphs(doc)
    if principal_photo and len(img_paras) > 0:
        _replace_image(doc, img_paras[0], principal_photo)
    if agent_photo and len(img_paras) > 1:
        _replace_image(doc, img_paras[1], agent_photo)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), total


def generate_multi_poa_bytes(
    template_bytes: bytes,
    case_info: dict,
    principals: List[dict],
    agent_photo: Optional[bytes] = None,
) -> tuple:
    """
    Generate a combined POA document for one or more principals.

    principals: list of dicts, each with per-principal fields plus an optional
                "photo" key (bytes) for that principal's ID image.

    Returns: (combined_docx_bytes, total_replacement_count)
    """
    doc_bytes_list = []
    total = 0

    for principal in principals:
        # Merge shared case fields with this principal's fields
        full_info = {**case_info, **{k: v for k, v in principal.items() if k != "photo"}}
        principal_photo = principal.get("photo")

        doc_bytes, count = generate_poa_bytes(
            template_bytes,
            full_info,
            principal_photo=principal_photo,
            agent_photo=agent_photo,
        )
        doc_bytes_list.append(doc_bytes)
        total += count

    combined = _merge_documents(doc_bytes_list)
    return combined, total


def get_csv_columns() -> list:
    return [f["key"] for f in FIELDS]


def get_example_row() -> dict:
    return {f["key"]: f["template_value"] for f in FIELDS}


def load_cases_from_csv(csv_bytes: bytes) -> tuple:
    text = csv_bytes.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    expected = set(get_csv_columns())
    cases, errors = [], []
    for i, row in enumerate(reader, start=2):
        clean = {k.strip(): v.strip() for k, v in row.items()}
        missing = expected - set(clean.keys())
        if missing:
            errors.append((i, f"Missing columns: {', '.join(sorted(missing))}"))
            continue
        errs = validate_case(clean)
        if errs:
            errors.append((i, "; ".join(errs)))
        else:
            cases.append(clean)
    return cases, errors
