"""
project_descriptions.py

Loads the boss's reference spreadsheet (company/project descriptions,
plus per-site/subfolder descriptions) and makes them available to
classify_project() and classify_address() as extra context, so the
model has real background on what each company does, typical email
domains, and keywords -- not just a bare folder name to match against.

Expected columns (same shape Numbers/Excel exports use -- header text
varies slightly release to release, so columns are matched by
POSITION, not by header text):
  A: ID proyecto   B: Empresa[\ Proyecto]   C: Dirección   D: Año
  E: Empresa\ Proyecto Descripción   F: Subfolder Descripción

Every sheet in the workbook is scanned (the boss adds one sheet per
year, in the same shape, and plans to eventually cover 2008-present)
-- rows are matched to real disk folders by their project CODE
("26-025", "26-003-01", ...), which is the one thing guaranteed to
line up between the spreadsheet and the real TRABAJOS folder names,
regardless of exact company/address text.

COST NOTE: growing this file to cover many more years does NOT
increase the token cost of any single classification. Matching is
scoped to the email's own year (see project_folders.py), so only that
year's candidates -- and therefore only that year's descriptions --
ever reach a prompt. Eighteen years of history can sit loaded in
memory for free; only ~1 year's worth of companies is ever "in play"
for a given email. What DOES matter for cost is how long each
individual description is, which is why they're truncated to
DESCRIPTION_MAX_CHARS below, cut at a sentence boundary when possible.

Where this goes: src/classification/project_descriptions.py.
"""

import os
import re
import openpyxl

from src.config import DESCRIPTIONS_XLSX_PATH, DESCRIPTION_MAX_CHARS

_CODE_RE = re.compile(r"^(\d{2}-\d+)")
_ADDRESS_PREFIX_RE = re.compile(r"^\d{2}-\d+-\d+\s+")

# Same US-code pattern used in src/output/plenergy_routing.py (kept as a
# separate copy here rather than imported, to keep this module
# self-contained -- it has no other dependency on the output package).
_US_CODE_RE = re.compile(r"US\s*-?\s*(\d{2,4}[A-Z]?)", re.IGNORECASE)

_cache = {"mtime": None, "company": {}, "address": {}, "us_code_index": {}}


def _extract_us_codes_from_text(text):
    if not text:
        return []
    return [f"US{m.group(1).upper()}" for m in _US_CODE_RE.finditer(str(text))]


def _normalize(text):
    return re.sub(r"\s+", " ", (text or "")).strip().upper()


def _truncate(text, max_chars):
    """Cuts a description down to a short summary instead of sending
    the full paragraph on every classification call. Tries to end on
    a sentence boundary (". ") within the limit so it doesn't cut off
    mid-word; falls back to a hard cut + "..." if there's no good
    sentence break available."""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    window = text[:max_chars]
    last_period = window.rfind(". ")
    if last_period > max_chars * 0.4:  # only use it if it's not too early
        return window[:last_period + 1]
    return window.rstrip() + "..."


def _load():
    """(Re)loads the workbook into memory if it's new or has changed
    on disk since last load. Cheap no-op on every call in between --
    safe to call before every single classification."""
    if not DESCRIPTIONS_XLSX_PATH or not os.path.exists(DESCRIPTIONS_XLSX_PATH):
        _cache["mtime"] = None
        _cache["company"] = {}
        _cache["address"] = {}
        _cache["us_code_index"] = {}
        return

    mtime = os.path.getmtime(DESCRIPTIONS_XLSX_PATH)
    if _cache["mtime"] == mtime:
        return  # already up to date

    company = {}
    address = {}  # address[company_code][normalized site text] = description
    us_code_index = {}  # us_code_index[normalized US code] = [project_id, ...]

    wb = openpyxl.load_workbook(DESCRIPTIONS_XLSX_PATH, data_only=True)
    for sheet in wb.worksheets:
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            continue
        for row in rows:
            if not row or not row[0]:
                continue
            project_id = str(row[0]).strip()
            site = row[2] if len(row) > 2 else None
            company_desc = row[4] if len(row) > 4 else None
            site_desc = row[5] if len(row) > 5 else None

            match = _CODE_RE.match(project_id)
            if not match:
                continue
            company_code = match.group(1)

            if company_desc and str(company_desc).strip() and company_code not in company:
                company[company_code] = _truncate(str(company_desc), DESCRIPTION_MAX_CHARS)

            if site_desc and str(site_desc).strip():
                address.setdefault(company_code, {})[_normalize(site)] = _truncate(
                    str(site_desc), DESCRIPTION_MAX_CHARS
                )

            # Index every US (Unidad de Suministro) code mentioned anywhere
            # in this row's own Dirección or Subfolder Descripción text,
            # against this row's own project_id (e.g. '25-003-06'). Used by
            # resolve_plenergy_folder's cross-year fallback to find a
            # station that was set up in an earlier year's TRABAJOS folder
            # -- possibly under different Plenergy branding (2025's
            # PLENOIL/DO PLENOIL vs 2026's PLENERGY/DO PLENERGY). Only
            # Plenergy-family rows ever mention a US code, so this index
            # stays small regardless of how many other companies/years are
            # in the sheet.
            for code in _extract_us_codes_from_text(site) + _extract_us_codes_from_text(site_desc):
                bucket = us_code_index.setdefault(code, [])
                if project_id not in bucket:
                    bucket.append(project_id)

    _cache["mtime"] = mtime
    _cache["company"] = company
    _cache["address"] = address
    _cache["us_code_index"] = us_code_index


def get_company_description(project_folder_name):
    """Looks up the company-level description for a real folder name
    like '26-025 PALMETILLO ALCALÁ DE GUADAÍRA', by its leading code.
    Returns None if there's no code, or no description on file."""
    _load()
    match = _CODE_RE.match(project_folder_name or "")
    if not match:
        return None
    return _cache["company"].get(match.group(1))


def get_address_description(company_folder_name, address_folder_name):
    """Looks up the site-level description for a specific address
    folder (e.g. '26-003-01 CAMÍ DE FAITANAR 2 - PICAÑA') under a
    given company folder, by matching the site's own name text
    (after stripping any address code prefix) against the
    spreadsheet's 'Dirección' column for that company's code."""
    _load()
    match = _CODE_RE.match(company_folder_name or "")
    if not match:
        return None
    site_map = _cache["address"].get(match.group(1))
    if not site_map:
        return None
    bare_address = _ADDRESS_PREFIX_RE.sub("", address_folder_name or "")
    return site_map.get(_normalize(bare_address))


def find_us_code_in_history(us_code):
    """Searches every year/sheet in the reference spreadsheet for a
    mention of the given US code (e.g. 'US584') in either the
    Dirección or Subfolder Descripción column. Used by
    resolve_plenergy_folder as a last-resort lookup when a station's US
    code isn't found among the CURRENT year's on-disk folders -- the
    station may have been set up in an earlier year's TRABAJOS folder
    (e.g. still under the 2025 PLENOIL/DO PLENOIL branding, rather than
    2026's PLENERGY/DO PLENERGY).

    Returns a list of project_id strings from the 'ID proyecto' column
    (e.g. ['25-003-06']) -- normally zero or one, but returned as a
    list in case the same code is genuinely mentioned on more than one
    row. Returns [] if the spreadsheet isn't configured, or doesn't
    mention this code anywhere."""
    _load()
    return list(_cache["us_code_index"].get(us_code, []))


def enrich_candidate_list(names, describe_fn):
    """Turns a flat list of folder names into a list of strings with
    their (truncated) description appended underneath (when one
    exists), ready to drop directly into an LLM prompt."""
    enriched = []
    for name in names:
        description = describe_fn(name)
        if description:
            enriched.append(f"{name}\n    Descripción: {description}")
        else:
            enriched.append(name)
    return enriched