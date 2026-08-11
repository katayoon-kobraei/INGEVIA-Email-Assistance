"""
eval_project_agent.py

Runs the real Gemini classifier (classify_project) against your actual
historical archive to measure matching accuracy -- how often it
correctly matches a real email to its real, already-known project
folder.

For every email found under TRABAJOS {year}\{project}\03.-CORREO\...,
this script:
  1. Builds the SAME candidate list the live pipeline would see --
     every real project folder that exists in that email's own year
     (matching is scoped to a single year, per the current design).
  2. Runs it through classify_project().
  3. Compares the predicted project name to the real one.

Every prediction (correct and wrong) is written to eval_results.xlsx
so you can review the misses and use them to tune project_prompt.md.

WARNING: this makes one real Gemini API call per email. For a large
archive, start with --limit to sample a subset before running everything.

ARCHIVE_ROOT comes from src/config.py (ARCHIVE_ROOT env var, falls
back to OUTPUT_ROOT) -- set it in .env rather than editing this file:
    ARCHIVE_ROOT=P:\

Setup (in your venv, with the project installed via `pip install -e .`
so `src` imports resolve):
    pip install extract-msg pandas openpyxl

Usage:
    python eval_project_agent.py                  # every email, default years
    python eval_project_agent.py --years 2026      # just one year
    python eval_project_agent.py --limit 30        # sample 30 emails first
"""

import argparse
import itertools
import os
import re

import extract_msg
import pandas as pd

from src.classification.project_agent import classify_project
from src.config import ARCHIVE_ROOT

# Written next to whichever ARCHIVE_ROOT is configured in .env.
OUTPUT_XLSX = os.path.join(ARCHIVE_ROOT, "eval_results.xlsx")


def find_project_folders(trabajos_root):
    """Direct children of TRABAJOS {year} -- the real, formal project
    folder names for that year. A historical archive has no holding
    pen; everything here is already a real project."""
    if not os.path.isdir(trabajos_root):
        return []
    return [
        name for name in os.listdir(trabajos_root)
        if os.path.isdir(os.path.join(trabajos_root, name))
    ]


def find_correo_folders(trabajos_root):
    """Recursively find every correspondence folder (any folder name
    containing "CORREO", e.g. "03.-CORREO", "2. CORREO" -- numbering
    isn't consistent project to project), at any nesting depth, and
    report which top-level project folder it belongs to."""
    for dirpath, dirnames, _ in os.walk(trabajos_root):
        if "CORREO" in os.path.basename(dirpath):
            rel = os.path.relpath(dirpath, trabajos_root)
            top_level_project = rel.split(os.sep)[0]
            yield top_level_project, dirpath
            dirnames[:] = []


def read_msg(file_path):
    try:
        msg = extract_msg.openMsg(file_path)
        data = {
            "subject": msg.subject or "",
            "sender": msg.sender or "",
            "to": msg.to or "",
            "body": (msg.body or "").strip(),
        }
        msg.close()
        return data
    except Exception:
        return None


def normalize(name):
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def iter_archive_emails(archive_root, years):
    """Yields (year, top_level_project, direction, conv_folder, msg_data)
    for every real historical email found."""
    for year in years:
        trabajos_folder = os.path.join(archive_root, f"TRABAJOS {year}")
        if not os.path.isdir(trabajos_folder):
            print(f"Skipping (not found): {trabajos_folder}")
            continue
        for top_level_project, correo_path in find_correo_folders(trabajos_folder):
            for direction in ("ENTRANTE", "SALIENTE"):
                direction_path = os.path.join(correo_path, direction)
                if not os.path.isdir(direction_path):
                    continue
                for conv_folder in sorted(os.listdir(direction_path)):
                    conv_path = os.path.join(direction_path, conv_folder)
                    if not os.path.isdir(conv_path):
                        continue
                    msg_files = [f for f in os.listdir(conv_path) if f.lower().endswith(".msg")]
                    if not msg_files:
                        continue
                    data = read_msg(os.path.join(conv_path, msg_files[0]))
                    if not data:
                        continue
                    yield year, top_level_project, direction, conv_folder, data


def run_eval(archive_root, years, limit=None):
    rows = []
    correct = 0
    total = 0

    projects_by_year = {
        year: find_project_folders(os.path.join(archive_root, f"TRABAJOS {year}"))
        for year in years
    }
    for year, names in projects_by_year.items():
        print(f"{year}: {len(names)} real project folder(s) on record")

    items = iter_archive_emails(archive_root, years)
    if limit:
        items = itertools.islice(items, limit)

    for year, top_level_project, direction, conv_folder, data in items:
        candidates = projects_by_year.get(year, [])

        fake_email = {
            "subject": data["subject"],
            "sender": data["sender"] if direction == "ENTRANTE" else None,
            "recipient": data["to"] if direction == "SALIENTE" else None,
            "direction": direction,
            "body": data["body"],
        }

        try:
            result = classify_project(fake_email, candidates)
            predicted_raw = result.project_folder_name
        except Exception as e:
            predicted_raw = f"ERROR: {e}"

        predicted = normalize(predicted_raw)
        actual = normalize(top_level_project)
        is_match = predicted == actual
        correct += is_match
        total += 1

        print(f"{'OK  ' if is_match else 'MISS'} | actual: {actual!r:45} | predicted: {predicted!r}")

        rows.append({
            "Year": year,
            "Actual Project": top_level_project,
            "Predicted": predicted_raw,
            "Correct": is_match,
            "Direction": direction,
            "Subject": data["subject"],
            "Conversation Folder": conv_folder,
        })

    if total:
        print(f"\nAccuracy: {correct}/{total} ({correct/total:.0%})")
    else:
        print("No emails evaluated -- check ARCHIVE_ROOT / --years.")

    if rows:
        pd.DataFrame(rows).to_excel(OUTPUT_XLSX, index=False, engine="openpyxl")
        print(f"Full results written to {OUTPUT_XLSX}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--years", nargs="+", type=int, default=[2026, 2025])
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()
    run_eval(ARCHIVE_ROOT, args.years, args.limit)