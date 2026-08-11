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
from src.output.project_folders import list_existing_addresses

DO_PLENERGY_FOLDER = "26-004 DO PLENERGY"
PLENERGY_FOLDER = "26-003 PLENERGY"

US_CODE_RE = re.compile(r"US\s*-?\s*(\d{2,4}[A-Z]?)", re.IGNORECASE)


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
    (project phase). Returns (project_folder_name, address_folder_name)
    if the station was found in either, or None if the US code doesn't
    match anything on file yet -- caller should fall back to the
    holding pen in that case."""
    do_match = find_address_by_us_code(output_root, company_year, DO_PLENERGY_FOLDER, us_code)
    if do_match:
        return DO_PLENERGY_FOLDER, do_match

    project_match = find_address_by_us_code(output_root, company_year, PLENERGY_FOLDER, us_code)
    if project_match:
        return PLENERGY_FOLDER, project_match

    return None