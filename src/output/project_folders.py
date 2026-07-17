import os
import re


def list_existing_projects(output_root, years):
    projects = []
    for year in years:
        trabajos_folder = os.path.join(output_root, f"TRABAJOS {year}")
        if os.path.exists(trabajos_folder):
            projects.extend(os.listdir(trabajos_folder))
    return sorted(projects)


def get_next_project_code(output_root, year):
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
    """Find every 'TRABAJOS {year}' folder that actually exists on disk,
    so the pipeline never has to hardcode / manually update a year list."""
    years = []
    if not os.path.isdir(output_root):
        return years
    for name in os.listdir(output_root):
        match = re.match(r"^TRABAJOS (\d{4})$", name)
        if match:
            years.append(int(match.group(1)))
    return sorted(years, reverse=True)


def get_project_year(project_folder_name):
    """Extract a project's creation year from its code prefix
    (e.g. '25-012 SOME CLIENT' -> 2025). Returns None if the name
    doesn't match the expected '{yy}-{seq} NAME' pattern."""
    match = re.match(r"^(\d{2})-\d+", project_folder_name)
    return 2000 + int(match.group(1)) if match else None