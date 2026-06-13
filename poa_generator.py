"""
poa_generator.py  —  Pure logic, no UI.
Handles field definitions, Word replacement, CSV parsing, and document generation.
"""

import io
import csv
from typing import Optional
from docx import Document


# ---------------------------------------------------------------------------
# Field definitions  (drives both the web form and the CSV column headers)
# ---------------------------------------------------------------------------
# Each dict describes one input field.
#   key            — internal identifier, also used as CSV column name
#   label          — human-readable label shown in the form
#   section        — groups fields together in the UI
#   template_value — the placeholder text that exists in the original template
#   required       — whether the field must be filled in
#   to_upper       — auto-convert to ALL CAPS
#   default        — pre-filled default value (empty string = no default)
#   help           — hint shown below the input box

FIELDS = [
    {
        "key":            "principal_name",
        "label":          "Principal Full Name",
        "section":        "Principal (Intended Parent)",
        "template_value": "CHENGFANG SHI",
        "required":       True,
        "to_upper":       True,
        "default":        "",
        "help":           "ALL CAPS — e.g., JOHN SMITH",
    },
    {
        "key":            "principal_role",
        "label":          "Principal Role",
        "section":        "Principal (Intended Parent)",
        "template_value": "Intended Father",
        "required":       False,
        "to_upper":       False,
        "default":        "Intended Father",
        "help":           "e.g., Intended Father or Intended Mother",
    },
    {
        "key":            "passport_country",
        "label":          "Principal Passport Country",
        "section":        "Principal (Intended Parent)",
        "template_value": "People’s Republic of China",   # smart apostrophe
        "required":       False,
        "to_upper":       False,
        "default":        "People’s Republic of China",
        "help":           "Country as it appears on passport",
    },
    {
        "key":            "principal_passport",
        "label":          "Principal Passport Number",
        "section":        "Principal (Intended Parent)",
        "template_value": "ED3297013",
        "required":       True,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., ED3297013",
    },
    {
        "key":            "child_last_name",
        "label":          "Child Last Name",
        "section":        "Child",
        "template_value": "SHI",          # used inside "CHILD SHI"
        "required":       True,
        "to_upper":       True,
        "default":        "",
        "help":           "Will appear as 'infant CHILD [LAST NAME]'",
    },
    {
        "key":            "surrogate_name",
        "label":          "Surrogate Full Name",
        "section":        "Surrogate",
        "template_value": "SELENA MARIA AISPURO",
        "required":       True,
        "to_upper":       True,
        "default":        "",
        "help":           "ALL CAPS — e.g., JANE DOE",
    },
    {
        "key":            "surrogate_dob",
        "label":          "Surrogate Date of Birth",
        "section":        "Surrogate",
        "template_value": "May 22, 1996",
        "required":       True,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., May 22, 1996",
    },
    {
        "key":            "due_date",
        "label":          "Due Date",
        "section":        "Birth Details",
        "template_value": "December 4, 2025",
        "required":       True,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., March 15, 2026",
    },
    {
        "key":            "hospital_name",
        "label":          "Hospital Name",
        "section":        "Birth Details",
        "template_value": "Loma Linda University Medical Center",
        "required":       False,
        "to_upper":       False,
        "default":        "Loma Linda University Medical Center",
        "help":           "Full hospital name",
    },
    {
        "key":            "hospital_city",
        "label":          "Hospital City",
        "section":        "Birth Details",
        "template_value": "Loma Linda",
        "required":       False,
        "to_upper":       False,
        "default":        "Loma Linda",
        "help":           "",
    },
    {
        "key":            "hospital_state",
        "label":          "Hospital State",
        "section":        "Birth Details",
        "template_value": "California",
        "required":       False,
        "to_upper":       False,
        "default":        "California",
        "help":           "",
    },
    {
        "key":            "agent_name",
        "label":          "Agent / Attorney-in-Fact Full Name",
        "section":        "Agent (Attorney-in-Fact)",
        "template_value": "JIAJIA GAO",
        "required":       True,
        "to_upper":       True,
        "default":        "",
        "help":           "ALL CAPS — person authorized to act on principal's behalf",
    },
    {
        "key":            "agent_dob",
        "label":          "Agent Date of Birth",
        "section":        "Agent (Attorney-in-Fact)",
        "template_value": "02/24/1986",
        "required":       True,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., 02/24/1986",
    },
    {
        "key":            "agent_passport",
        "label":          "Agent Passport Number",
        "section":        "Agent (Attorney-in-Fact)",
        "template_value": "EJ4979964",
        "required":       True,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., EJ4979964",
    },
    {
        "key":            "attorney_name",
        "label":          "Handling Attorney",
        "section":        "Firm",
        "template_value": "Xuelan Fang",
        "required":       False,
        "to_upper":       False,
        "default":        "Xuelan Fang",
        "help":           "Attorney at Tsong Law Group handling this case",
    },
    {
        "key":            "agency_name",
        "label":          "Fertility Agency (optional, used in filename)",
        "section":        "Firm",
        "template_value": "",
        "required":       False,
        "to_upper":       False,
        "default":        "",
        "help":           "e.g., C&T Fertility Consultant",
    },
]

# Replacement order matters: specific strings before the substrings they contain.
# "Loma Linda University Medical Center" must be replaced before plain "Loma Linda".
REPLACEMENT_ORDER = [
    "hospital_name",
    "hospital_city",
    "hospital_state",
    "due_date",
    "surrogate_dob",
    "agent_dob",
    "principal_name",
    "child_last_name",   # handled specially — replaces "CHILD SHI" not just "SHI"
    "surrogate_name",
    "agent_name",
    "attorney_name",
    "principal_passport",
    "agent_passport",
    "passport_country",
    "principal_role",
]


# ---------------------------------------------------------------------------
# Low-level Word replacement engine
# ---------------------------------------------------------------------------
# Word stores paragraph text split across many tiny "runs" in its XML.
# A single word like "02/24/1986" may be 7 separate fragments.
# These functions stitch them back together for reliable find-and-replace.

def _replace_once(para, old: str, new: str) -> bool:
    """Replace first occurrence of old across fragmented runs. Returns True if replaced."""
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
# Public API
# ---------------------------------------------------------------------------

def validate_case(info: dict) -> list:
    """Return a list of error strings. Empty list means the case is valid."""
    errors = []
    for field in FIELDS:
        if field["required"] and not info.get(field["key"], "").strip():
            errors.append(f"'{field['label']}' is required.")
    return errors


def make_filename(info: dict) -> str:
    """Build a descriptive filename for the generated document."""
    principal_last = info.get("principal_name", "Client").split()[-1]
    surrogate_last = info.get("surrogate_name", "Surrogate").split()[-1]
    agency = info.get("agency_name", "").strip()
    parts = [f"POA - {principal_last} & {surrogate_last}"]
    if agency:
        parts.append(agency)
    return " - ".join(parts) + ".docx"


def generate_poa_bytes(template_bytes: bytes, info: dict) -> tuple:
    """
    Generate a filled POA document.

    Args:
        template_bytes: The raw bytes of the template .docx file.
        info: Dict mapping field keys to user-provided values.

    Returns:
        (docx_bytes, replacement_count)
    """
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
            # Replace "CHILD SHI" (the full pattern), not just "SHI"
            old = f"CHILD {field['template_value']}"
            new = f"CHILD {value}"
            total += _replace_in_doc(doc, old, new)
        else:
            total += _replace_in_doc(doc, template_val, value)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), total


def get_csv_columns() -> list:
    """Return CSV column names in the correct order (one per field)."""
    return [f["key"] for f in FIELDS]


def get_empty_row() -> dict:
    """Return a dict with all field keys set to their default values."""
    return {f["key"]: f["default"] for f in FIELDS}


def get_example_row() -> dict:
    """Return the Shi/Aispuro case as an example CSV row."""
    return {f["key"]: f["template_value"] for f in FIELDS}


def load_cases_from_csv(csv_bytes: bytes) -> tuple:
    """
    Parse a CSV file into a list of case dicts.

    Returns:
        (cases, errors)
        cases  — list of dicts, one per valid row
        errors — list of (row_number, message) for bad rows
    """
    text = csv_bytes.decode("utf-8-sig")   # utf-8-sig handles Excel's BOM
    reader = csv.DictReader(io.StringIO(text))

    expected_columns = set(get_csv_columns())
    cases, errors = [], []

    for i, row in enumerate(reader, start=2):   # row 1 = header
        # Strip whitespace from all values
        clean = {k.strip(): v.strip() for k, v in row.items()}

        # Check for unexpected columns (just warn, don't fail)
        missing = expected_columns - set(clean.keys())
        if missing:
            errors.append((i, f"Missing columns: {', '.join(sorted(missing))}"))
            continue

        row_errors = validate_case(clean)
        if row_errors:
            errors.append((i, "; ".join(row_errors)))
        else:
            cases.append(clean)

    return cases, errors
