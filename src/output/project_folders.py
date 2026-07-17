import os
import re

RESERVED_TOP_LEVEL_NAMES = {"UNSORTED"}


def get_holding_pen_name(year):
    """The per-year holding folder for emails that don't match any
    existing project and aren't a started/formal project yet."""
    yy = year % 100
    return f"{yy:02d}-000 MAILS"


def is_formal_project_code(project_folder_name):
    """True if the name already has a real sequential project code
    prefix (e.g. '26-011 MEETPACK BENASAU')."""
    return re.match(r"^\d{2}-\d+", project_folder_name) is not None


def list_existing_projects(output_root, years):
    """Flat list of matchable candidate names for the given year(s):
    real top-level project folders, plus every not-yet-started entry
    sitting inside that year's holding pen (e.g. 26-000 MAILS).
    UNSORTED is excluded -- it's not a real matchable project, just
    where failed classifications get dumped."""
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
                if os.path.isdir(pen_path):
                    projects.extend(os.listdir(pen_path))
            else:
                projects.append(name)
    return sorted(projects)


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


def resolve_project_relative_path(output_root, project_folder_name, email_year):
    """Decide where a project's folder actually lives.

    Returns (year, relative_path_under_TRABAJOS_year):
      - Formal projects (have a '{yy}-{seq}' code) and UNSORTED stay
        exactly where they've always lived, at the top level, in the
        project's own year (from its code).
      - Everything else (a bare descriptive name, no code -- i.e. a
        not-yet-started project) is filed under the EMAIL's own year,
        inside that year's holding pen ('{yy}-000 MAILS/{name}').
        Matching is scoped to a single year (see list_existing_projects),
        so a bare name is always treated as belonging to the year it
        was just seen in -- it will not be merged with a same-named
        entry from a different year.
    """
    if is_formal_project_code(project_folder_name) or project_folder_name in RESERVED_TOP_LEVEL_NAMES:
        year = get_project_year(project_folder_name) or email_year
        return year, project_folder_name

    relative_path = os.path.join(get_holding_pen_name(email_year), project_folder_name)
    return email_year, relative_path


def list_existing_addresses(output_root, company_year, company_folder_name):
    """Existing address subfolders for one specific formal company
    (e.g. inside '26-003 PLENERGY'), full names including their own
    '{company_code}-{NN}' prefix. Only meaningful for a company that
    already has a real project code -- returns [] otherwise."""
    company_path = os.path.join(output_root, f"TRABAJOS {company_year}", company_folder_name)
    if not os.path.isdir(company_path):
        return []
    company_code_match = re.match(r"^(\d{2}-\d+)", company_folder_name)
    if not company_code_match:
        return []
    company_code = company_code_match.group(1)
    pattern = re.compile(rf"^{re.escape(company_code)}-\d+\s")
    return sorted(
        name for name in os.listdir(company_path)
        if os.path.isdir(os.path.join(company_path, name)) and pattern.match(name)
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