"""
plenergy_routing.py

Special-case routing for the PLENERGY / DO PLENERGY relationship. The
same physical station ("US" -- Unidad de Suministro) can live under
either folder depending on its phase:

  26-003 PLENERGY     -- pre-construction (project, licensing, due diligence)
  26-004 DO PLENERGY  -- construction phase (Direccion de Obra)

Rather than leaving classify_project to guess which phase folder an
email belongs to and then classify_address to match/mint an address
within just that one company, this does a deterministic cross-folder
lookup keyed by the station's own US code -- the one identifier
guaranteed to be consistent across both folders for the same physical
station. DO PLENERGY is checked first (construction correspondence
about a station takes priority once it's under construction), then
PLENERGY. If the US code doesn't match anything in either folder yet,
the caller falls back to the holding pen.

Where this goes: src/output/plenergy_routing.py

KNOWN LIMITATION: this only searches the CURRENT email's year. Last
year's folders used different branding for the same client (25-003
PLENOIL / 25-004 DO PLENOIL, not PLENERGY) -- a station opened in 2025
that's still getting correspondence won't be found by this lookup. Ask
if you want cross-year / cross-naming support added.
"""

import re

from src.config import PLENERGY_SENDER_DOMAINS
from src.output.project_folders import list_existing_addresses, list_existing_companies
from src.classification.project_descriptions import find_us_code_in_history

DO_PLENERGY_FOLDER = "26-004 DO PLENERGY"
PLENERGY_FOLDER = "26-003 PLENERGY"

US_CODE_RE = re.compile(r"US\s*-?\s*(\d{2,4}[A-Z]?)", re.IGNORECASE)
_COMPANY_CODE_RE = re.compile(r"^(\d{2})-\d+")


def is_plenergy_sender(email):
    """True if the sender's domain is Plenergy's own (plenergy.es) or
    their technical intermediary PLAINCO's (plainco.es) -- both
    folders' descriptions cite these as their typical senders."""
    sender = (email.get("sender") or "").strip().lower()
    if "@" not in sender:
        return False
    domain = sender.rsplit("@", 1)[-1]
    domains = [d.strip().lower() for d in PLENERGY_SENDER_DOMAINS.split(",") if d.strip()]
    return any(domain == d or domain.endswith("." + d) for d in domains)


def extract_us_codes(email):
    """Pulls every distinct 'US' (Unidad de Suministro) code out of the
    email's subject or body -- e.g. 'US552', 'US188A' -- tolerating the
    usual formatting variations ('US 552', '(US574)'). Returns them as
    a list of normalized 'US{code}' strings, in the order they first
    appear, deduplicated. Returns an empty list if none are found.

    Some emails mention two stations at once (e.g. a subject naming
    both US552 and US574) -- those need to be filed under both
    stations' folders, not just the first one found, which is why this
    returns every match instead of just the first."""
    text = f"{email.get('subject') or ''} {email.get('body') or ''}"
    codes = []
    for match in US_CODE_RE.finditer(text):
        code = f"US{match.group(1).upper()}"
        if code not in codes:
            codes.append(code)
    return codes


def extract_us_code(email):
    """Same as extract_us_codes, but returns just the first code found
    (or None) -- kept for any caller that only ever expects a single
    station per email."""
    codes = extract_us_codes(email)
    return codes[0] if codes else None


def _normalize_for_us_match(text):
    return re.sub(r"US\s*-?\s*", "US", text.upper())


def find_address_by_us_code(output_root, company_year, company_folder_name, us_code):
    """Deterministic check: does this company already have an address
    subfolder whose name contains this exact US code? Returns the
    matching folder name, or None."""
    for name in list_existing_addresses(output_root, company_year, company_folder_name):
        if us_code in _normalize_for_us_match(name):
            return name
    return None


def resolve_plenergy_folder(output_root, company_year, us_code):
    """Cross-folder lookup for a Plenergy-family email, in priority
    order: DO PLENERGY (construction phase) first, then PLENERGY
    (project phase), both within the email's OWN year. Returns
    (project_folder_name, address_folder_name) if the station was
    found, or None if the US code doesn't match anything on file yet
    -- caller should fall back to the holding pen in that case.

    If nothing matches in the email's own year, also checks the boss's
    reference spreadsheet (DESCRIPTIONS_XLSX_PATH, via
    project_descriptions.find_us_code_in_history) for this US code
    under any OTHER year -- a station can have been set up in an
    earlier year's TRABAJOS folder, possibly under different Plenergy
    branding (2025's PLENOIL/DO PLENOIL vs 2026's PLENERGY/DO
    PLENERGY). The spreadsheet is used only to discover WHICH year and
    company code to look at -- the folder name actually returned always
    comes from the real, on-disk company/address lookup for that year,
    never from the spreadsheet's own text (which is sometimes
    incomplete -- see find_us_code_in_history)."""
    do_match = find_address_by_us_code(output_root, company_year, DO_PLENERGY_FOLDER, us_code)
    if do_match:
        return DO_PLENERGY_FOLDER, do_match

    project_match = find_address_by_us_code(output_root, company_year, PLENERGY_FOLDER, us_code)
    if project_match:
        return PLENERGY_FOLDER, project_match

    return _resolve_plenergy_folder_from_history(output_root, us_code)


def _find_company_folder_by_code(existing_companies, company_code):
    """Finds the real on-disk company folder for a given code prefix
    (e.g. '25-003'), regardless of what the company was actually named
    that year -- Plenergy's own branding has changed year to year
    (PLENOIL/DO PLENOIL in 2025, PLENERGY/DO PLENERGY from 2026), so
    this matches by CODE, never by assuming a fixed company name."""
    prefix = f"{company_code} ".casefold()
    for name in existing_companies:
        if str(name).strip().casefold().startswith(prefix):
            return name
    return None


def _resolve_plenergy_folder_from_history(output_root, us_code):
    """The cross-year half of resolve_plenergy_folder -- see its
    docstring. Tries every project_id the spreadsheet associates with
    this US code, in whichever years those actually turn out to be.

    A station can genuinely be listed twice in the same historical
    year -- once under its project-phase entry, once under its
    construction-phase ("DO ...") entry, exactly like the same
    duplication the current-year lookup already resolves by checking
    DO_PLENERGY_FOLDER first. So collect every real on-disk match
    first, then prefer a "DO ..." company folder over a non-"DO" one,
    instead of just taking whichever the spreadsheet happens to list
    first."""
    candidates = []  # (is_do_phase, company_folder_name, address_match)
    for project_id in find_us_code_in_history(us_code):
        match = _COMPANY_CODE_RE.match(project_id)
        if not match:
            continue
        yy, company_code = match.group(1), match.group(0)
        historical_year = 2000 + int(yy)

        existing_companies = list_existing_companies(output_root, historical_year)
        company_folder_name = _find_company_folder_by_code(existing_companies, company_code)
        if not company_folder_name:
            continue  # spreadsheet mentions it, but no matching folder exists on disk

        address_match = find_address_by_us_code(output_root, historical_year, company_folder_name, us_code)
        if not address_match:
            continue

        company_label = re.sub(r"^\d{2}-\d+\s+", "", company_folder_name).strip().upper()
        is_do_phase = company_label.startswith("DO ")
        candidates.append((is_do_phase, company_folder_name, address_match))

    if not candidates:
        return None

    candidates.sort(key=lambda c: not c[0])  # DO-phase entries first
    _, company_folder_name, address_match = candidates[0]
    return company_folder_name, address_match