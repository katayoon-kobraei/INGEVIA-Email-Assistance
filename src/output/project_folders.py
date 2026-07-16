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