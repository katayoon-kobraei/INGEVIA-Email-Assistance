import json
import os
import re

RESERVED_TOP_LEVEL_NAMES = {"UNSORTED"}


def find_correo_folder(path):
    """Returns the name of the correspondence folder directly under
    `path`, if one exists -- any immediate subfolder whose name
    contains "CORREO" (exact capitalization), regardless of its
    numbering prefix. Different companies on the server use different
    conventions ("03.-CORREO", "2. CORREO", "CORREO", ...), so this
    matches by substring rather than assuming a fixed "03.-CORREO"
    name. Returns None if `path` doesn't exist or has no such
    subfolder. If more than one matches (shouldn't normally happen),
    returns the first in directory-listing order."""
    if not os.path.isdir(path):
        return None
    for name in os.listdir(path):
        if "CORREO" in name and os.path.isdir(os.path.join(path, name)):
            return name
    return None


def company_uses_address_subfolders(output_root, company_year, company_folder_name):
    """Deterministically decides whether THIS company already uses
    address-level subfoldering, by looking at its real folder
    structure -- never left to the model's judgment about what a
    given email's text happens to mention.

    Returns:
      True  -- company already has one or more address-coded
               subfolders (e.g. '26-003-01 ...'). classify_address
               MUST run, regardless of this email's own content.
      False -- company already has a correspondence folder (any name
               containing "CORREO", e.g. "03.-CORREO" or "2. CORREO")
               directly under it (single-site). classify_address must
               NOT run for it.
      None  -- company folder doesn't exist yet, or exists but has
               neither a CORREO folder nor any address subfolder yet
               (brand new). Ambiguous -- caller should fall back to
               the email's own content (mentions_specific_address)
               to decide how to set this company up for the first time.
    """
    company_path = get_company_path(output_root, company_year, company_folder_name)
    if not os.path.isdir(company_path):
        return None

    if find_correo_folder(company_path):
        return False
    if list_existing_addresses(output_root, company_year, company_folder_name):
        return True
    return None


def find_existing_holding_pen_entry(output_root, year, bare_name):
    """Checks this year's holding pen for an entry matching bare_name
    (ignoring any '{NNN} ' counter prefix it may already have) -- so
    repeat unmatched emails for the same special case land in the SAME
    holding-pen folder instead of minting a new one every time."""
    pen_path = os.path.join(output_root, f"TRABAJOS {year}", get_holding_pen_name(year))
    if not os.path.isdir(pen_path):
        return None
    target = bare_name.strip().upper()
    for name in os.listdir(pen_path):
        stripped = re.sub(r"^\d{3}\s+", "", name).strip().upper()
        if stripped == target:
            return name
    return None

def get_holding_pen_name(year):
    """The per-year holding folder for emails that don't match any
    existing project and aren't a started/formal project yet."""
    yy = year % 100
    return f"{yy:02d}-000 MAILS"


def is_formal_project_code(project_folder_name):
    """True if the name already has a real sequential project code
    prefix (e.g. '26-011 MEETPACK BENASAU')."""
    return re.match(r"^\d{2}-\d+", project_folder_name) is not None


def _distinct_holding_pen_projects(pen_path):
    """Each entry in a year's holding pen is now a flat, one-per-email
    folder (see build_holding_pen_folder_name) -- its own display name
    bakes in the date/time/sender/subject, so it's no longer a clean
    matchable project name on its own. The real project name is read
    back out of each entry's metadata.json instead (written by
    save_email/save_plenergy_fallback_email). Plenergy fallback entries
    have no "project_folder" key at all and are correctly skipped."""
    names = set()
    if not os.path.isdir(pen_path):
        return names
    for entry in os.listdir(pen_path):
        meta_path = os.path.join(pen_path, entry, "metadata.json")
        if not os.path.isfile(meta_path):
            continue
        try:
            with open(meta_path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue
        project = data.get("project_folder")
        if project:
            names.add(project)
    return names



def list_existing_companies(output_root, year):
    """Return the real top-level company/client folders for one TRABAJOS year.

    The holding pen (``YY-000 MAILS``) and ``UNSORTED`` are operational
    buckets, not companies, so they are never offered to the company matcher.
    This is intentionally filesystem-driven: the AI may only select a company
    that actually exists on the server.
    """
    trabajos_folder = os.path.join(output_root, f"TRABAJOS {year}")
    if not os.path.isdir(trabajos_folder):
        return []

    excluded = {get_holding_pen_name(year), *RESERVED_TOP_LEVEL_NAMES}
    return sorted(
        name for name in os.listdir(trabajos_folder)
        if name not in excluded
        and not name.startswith(".")
        and os.path.isdir(os.path.join(trabajos_folder, name))
    )


def exact_existing_name(candidate, existing_names):
    """Resolve an AI-returned folder name against a real candidate list.

    Matching is case-insensitive but otherwise exact. This prevents a model
    hallucination or slightly rewritten name from ever creating a fake company
    or project folder. Returns the real on-disk name, or ``None``.
    """
    target = str(candidate or "").strip().casefold()
    if not target:
        return None
    for name in existing_names:
        if str(name).strip().casefold() == target:
            return name
    return None

def list_existing_projects(output_root, years):
    """Flat list of matchable candidate names for the given year(s):
    real top-level project folders, plus every distinct not-yet-formal
    project name already seen in that year's holding pen (e.g.
    26-000 MAILS). UNSORTED is excluded -- it's not a real matchable
    project, just where failed classifications get dumped."""
    projects = []
    for year in years:
        trabajos_folder = os.path.join(output_root, f"TRABAJOS {year}")
        if not os.path.exists(trabajos_folder):
            continue
        pen_name = get_holding_pen_name(year)
        for name in os.listdir(trabajos_folder):
            if name == "UNSORTED":
                continue
            if name == pen_name:
                pen_path = os.path.join(trabajos_folder, name)
                projects.extend(_distinct_holding_pen_projects(pen_path))
            else:
                projects.append(name)
    return sorted(set(projects))


def get_next_project_code(output_root, year):
    """Only relevant when a human/tool formally promotes something to
    a real numbered project. The automated pipeline no longer calls
    this on its own -- unmatched emails go to that year's holding pen
    instead of getting an automatic new code."""
    trabajos_folder = os.path.join(output_root, f"TRABAJOS {year}")
    yy = year % 100
    pattern = re.compile(rf"^{yy:02d}-(\d+)")
    max_seq = 0
    if os.path.exists(trabajos_folder):
        for name in os.listdir(trabajos_folder):
            match = pattern.match(name)
            if match:
                max_seq = max(max_seq, int(match.group(1)))
    return f"{yy:02d}-{max_seq + 1:03d}"


def discover_trabajos_years(output_root):
    """Find every 'TRABAJOS {year}' folder that actually exists on disk.
    Not used for matching anymore (matching is scoped to a single
    year), but still handy for admin/reporting tools that want to see
    every year at once."""
    years = []
    if not os.path.isdir(output_root):
        return years
    for name in os.listdir(output_root):
        match = re.match(r"^TRABAJOS (\d{4})$", name)
        if match:
            years.append(int(match.group(1)))
    return sorted(years, reverse=True)


def get_project_year(project_folder_name):
    """Extract a formal project's creation year from its code prefix
    (e.g. '25-012 SOME CLIENT' -> 2025). Returns None if the name
    doesn't match the expected '{yy}-{seq} NAME' pattern (e.g. a bare
    holding-pen name, or UNSORTED)."""
    match = re.match(r"^(\d{2})-\d+", project_folder_name)
    return 2000 + int(match.group(1)) if match else None


def has_holding_pen_counter(name):
    """True if a holding-pen entry name already has its local '{NNN}'
    counter prefix (e.g. '001 PLENERGY-PLAINCO')."""
    return re.match(r"^\d{3}\s", name) is not None


def get_next_holding_pen_counter(output_root, year):
    """Next sequential local counter (e.g. '004') for a brand-new entry
    in that year's holding pen ('{yy}-000 MAILS'). This is purely a
    local, human-friendly ordering within the pen -- unrelated to, and
    replaced by, the real project code if/when a project is formally
    promoted out of the pen."""
    pen_path = os.path.join(output_root, f"TRABAJOS {year}", get_holding_pen_name(year))
    max_seq = 0
    if os.path.isdir(pen_path):
        for name in os.listdir(pen_path):
            match = re.match(r"^(\d{3})\s", name)
            if match:
                max_seq = max(max_seq, int(match.group(1)))
    return f"{max_seq + 1:03d}"


def resolve_project_relative_path(output_root, project_folder_name, email_year):
    """Decide where a project's folder actually lives.

    Returns (year, relative_path_under_TRABAJOS_year):
      - Formal projects (have a '{yy}-{seq}' code) and UNSORTED stay
        exactly where they've always lived, at the top level, in the
        project's own year (from its code).
      - Everything else (a bare descriptive name, no code -- i.e. a
        not-yet-started project) is filed under the EMAIL's own year,
        inside that year's holding pen ('{yy}-000 MAILS/{NNN name}').
        A brand-new bare name (no matching entry yet, so it has no
        '{NNN}' counter) gets the pen's next local counter minted for
        it automatically here -- the model itself never assigns this,
        same as it never assigns a real project code. Matching is
        scoped to a single year (see list_existing_projects), so a
        bare name is always treated as belonging to the year it was
        just seen in -- it will not be merged with a same-named entry
        from a different year.
    """
    if is_formal_project_code(project_folder_name) or project_folder_name in RESERVED_TOP_LEVEL_NAMES:
        year = get_project_year(project_folder_name) or email_year
        return year, project_folder_name

    pen_entry_name = project_folder_name
    if not has_holding_pen_counter(pen_entry_name):
        counter = get_next_holding_pen_counter(output_root, email_year)
        pen_entry_name = f"{counter} {pen_entry_name}"

    relative_path = os.path.join(get_holding_pen_name(email_year), pen_entry_name)
    return email_year, relative_path


def get_company_path(output_root, company_year, company_folder_name):
    """Resolve where a company's folder actually lives on disk, whether
    it's a formal coded project (top-level) or a bare not-yet-formal
    name sitting inside that year's holding pen."""
    if is_formal_project_code(company_folder_name) or company_folder_name in RESERVED_TOP_LEVEL_NAMES:
        return os.path.join(output_root, f"TRABAJOS {company_year}", company_folder_name)
    return os.path.join(
        output_root, f"TRABAJOS {company_year}", get_holding_pen_name(company_year), company_folder_name
    )


def list_existing_addresses(output_root, company_year, company_folder_name):
    """Existing address subfolders for one specific company. For a
    formal company (e.g. '26-003 PLENERGY'), returns full names
    including their own '{company_code}-{NN}' prefix. For a
    not-yet-formal (holding-pen) company, there's no code sequence yet
    -- returns the bare address names directly instead."""
    company_path = get_company_path(output_root, company_year, company_folder_name)
    if not os.path.isdir(company_path):
        return []
    company_code_match = re.match(r"^(\d{2}-\d+)", company_folder_name)
    if company_code_match:
        company_code = company_code_match.group(1)
        pattern = re.compile(rf"^{re.escape(company_code)}-\d+\s")
        return sorted(
            name for name in os.listdir(company_path)
            if os.path.isdir(os.path.join(company_path, name)) and pattern.match(name)
        )
    # Holding-pen company: no code sequence exists yet, so address
    # subfolders are just bare names. The correspondence folder itself
    # (any name containing "CORREO", e.g. "03.-CORREO" or "2. CORREO")
    # lives directly under the company folder for emails with no
    # specific site, so it must be excluded from the address candidate
    # list.
    return sorted(
        name for name in os.listdir(company_path)
        if os.path.isdir(os.path.join(company_path, name)) and "CORREO" not in name
    )


def get_next_address_code(output_root, company_year, company_folder_name):
    """Next sequential address code for a formal company, e.g.
    '26-003-04' if '26-003-01'..'26-003-03' already exist. Unlike
    get_next_project_code, this is safe to call automatically -- the
    company itself was already a vetted, real client; assigning it
    another site under the same client is a low-risk, mechanical
    decision, not a new-client decision."""
    company_code_match = re.match(r"^(\d{2})-(\d+)", company_folder_name)
    if not company_code_match:
        raise ValueError(f"'{company_folder_name}' has no formal project code")
    company_code = f"{company_code_match.group(1)}-{company_code_match.group(2)}"
    pattern = re.compile(rf"^{re.escape(company_code)}-(\d+)")
    max_addr_seq = 0
    for name in list_existing_addresses(output_root, company_year, company_folder_name):
        match = pattern.match(name)
        if match:
            max_addr_seq = max(max_addr_seq, int(match.group(1)))
    return f"{company_code}-{max_addr_seq + 1:02d}"