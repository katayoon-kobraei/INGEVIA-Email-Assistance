"""
reprocess_unsorted.py

Standalone maintenance tool: re-runs classification on every email already
sitting in <ARCHIVE_ROOT>/TRABAJOS <year>/UNSORTED (the current year's
catch-all folder for emails the pipeline could not classify -- most often
because a Gemini API call failed or disconnected mid-run), WITHOUT ever
opening Outlook. Safe to run any time, independently of the scheduled
pipeline task -- it only reads the already-saved email.txt/metadata.json
files sitting on disk and talks to Gemini.

For every entry found there, this:
  1. Re-runs the same classification cascade as src/pipeline.py (US-code
     match -> Plenergy sender match -> company/project match -> Plenergy
     rescue check), using the entry's already-saved data.
  2. If it now resolves to a real company/project (or a Plenergy station,
     or at least a real company with no project match), MOVES the
     already-saved folder to its correct destination -- computed by the
     exact same resolve_save_destination() that save_email() uses -- and
     updates (never duplicates) its row in index.csv and
     email_report_log.csv, then regenerates the status page and
     "Informe de Emails.xlsx" so everything reflects it immediately.
  3. If Gemini judges it irrelevant (social media, marketing/newsletters,
     automated receipts/bounces, spam, personal mail), DELETES the folder
     and its index/report rows permanently.
  4. Otherwise (classification itself failed again -- Gemini still
     unreachable) leaves it untouched in UNSORTED for a future run.

Run manually from the app folder:
    python reprocess_unsorted.py            (moves/deletes for real)
    python reprocess_unsorted.py --dry-run   (only prints the plan)

--dry-run is strongly recommended for a first pass so you can see exactly
what would move and what would be deleted before anything actually
happens.

NOTE on classification logic: the cascade below intentionally mirrors
src/pipeline.py's per-email logic (same functions, same order, same
fallback rules) but is its own copy here -- the same pattern already used
elsewhere in this codebase (see the _US_CODE_RE comment in
src/classification/project_descriptions.py) so this standalone tool has
no import-time dependency on pipeline.py's Outlook-specific code. If you
ever change the classification cascade in src/pipeline.py, mirror the
same change here.

NOTE on data limits: an UNSORTED entry's email.txt/metadata.json never
recorded the original email's full To/Cc header text (only Outlook had
that, live) -- so the billing-forwarding check pipeline.py normally runs
first (classify_billing + boss_is_recipient) is skipped here; it needs
data that was never saved to disk. A misfiled billing email will still
get correctly filed once its company/project resolves through the normal
cascade below, just without the automatic forward-to-administracion step.
"""

from __future__ import annotations

import argparse
import csv
import datetime
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
os.chdir(APP_ROOT)
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from src.config import ARCHIVE_ROOT, OUTPUT_ROOT, ensure_output_root
from src.output.billing_routing import is_internal_sender, is_ignored_sender
from src.output.plenergy_routing import (
    is_plenergy_sender, extract_us_codes, resolve_plenergy_folder,
    DO_PLENERGY_FOLDER, PLENERGY_FOLDER,
)
from src.classification.plenergy_agent import classify_plenergy_address
from src.output.project_folders import (
    list_existing_companies, list_existing_addresses, get_project_year,
    get_holding_pen_name, exact_existing_name,
)
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.save_email import resolve_save_destination
from src.output.report_writer import cheap_fallback_summary, generate_email_report_xlsx, REPORT_FIELDNAMES
from src.output.index_writer import FIELDNAMES as INDEX_FIELDNAMES
from src.output.status_page import generate_status_page

UNSORTED_FOLDER_NAME = "UNSORTED"


# ---------------------------------------------------------------------------
# Finding and parsing already-saved entries
# ---------------------------------------------------------------------------

def _unsorted_root(year: int) -> str:
    return os.path.join(ARCHIVE_ROOT, f"TRABAJOS {year}", UNSORTED_FOLDER_NAME)


def _find_entries(unsorted_root: str):
    """Yields the path of every email entry folder under UNSORTED, at any
    nesting depth -- handles both a flat one-folder-per-email layout (like
    the YY-000 MAILS holding pen) and a nested COMPANY/CORREO/DIRECTION
    layout, since it isn't known in advance which shape produced entries
    here. An "entry" is any directory that directly contains an
    email.txt file (written by every save_* function in save_email.py)."""
    for dirpath, dirnames, filenames in os.walk(unsorted_root):
        if "email.txt" in filenames:
            yield dirpath
            dirnames[:] = []  # don't descend into an entry's own subfolders


def _parse_timestamp(value):
    if not value:
        return None
    if isinstance(value, datetime.datetime):
        return value
    text = str(value).strip()
    for candidate in (text, text.replace(" ", "T", 1)):
        try:
            return datetime.datetime.fromisoformat(candidate)
        except ValueError:
            continue
    for fmt in ("%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _parse_email_txt(raw: str):
    """email.txt is always either 'From: ...\\n' or 'To: ...\\n', then
    'Subject: ...\\nDate: ...\\n\\n', then the body -- see save_email.py."""
    header, _, body = raw.partition("\n\n")
    header_map = {}
    for line in header.splitlines():
        if ":" in line:
            key, _, value = line.partition(":")
            header_map[key.strip().lower()] = value.strip()
    return header_map, body


def load_entry(entry_path: str):
    """Reconstructs an email-like dict from an already-saved entry folder:
    metadata.json is authoritative when present, email.txt's own header
    fills any gaps. Returns None if not enough information survives to
    classify it at all (missing subject/timestamp/direction)."""
    meta = {}
    meta_path = os.path.join(entry_path, "metadata.json")
    if os.path.isfile(meta_path):
        try:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
        except (OSError, json.JSONDecodeError):
            meta = {}

    txt_path = os.path.join(entry_path, "email.txt")
    try:
        with open(txt_path, encoding="utf-8") as f:
            raw = f.read()
    except OSError:
        return None
    header_map, body = _parse_email_txt(raw)

    direction = str(meta.get("direction") or "").strip().upper()
    if direction not in ("ENTRANTE", "SALIENTE"):
        direction = "SALIENTE" if "to" in header_map and "from" not in header_map else "ENTRANTE"

    timestamp = _parse_timestamp(meta.get("timestamp")) or _parse_timestamp(header_map.get("date"))
    if timestamp is None:
        return None
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=datetime.timezone.utc)

    subject = str(meta.get("subject") or header_map.get("subject") or "").strip() or "(Sin asunto)"
    sender = meta.get("sender") or (header_map.get("from") if direction == "ENTRANTE" else None)
    recipient = meta.get("recipient") or (header_map.get("to") if direction == "SALIENTE" else None)
    entry_id = str(meta.get("id") or "").strip()

    return {
        "id": entry_id,
        "store_id": meta.get("store_id") or "",
        "subject": subject,
        "body": body,
        "direction": direction,
        "sender": sender,
        "recipient": recipient,
        "sender_name": meta.get("sender_name") or "",
        "sender_email": meta.get("sender_email") or sender or "",
        "timestamp": timestamp,
        "_entry_path": entry_path,
        "_metadata": meta,
        "_metadata_path": meta_path if os.path.isfile(meta_path) else None,
    }


# ---------------------------------------------------------------------------
# Classification cascade -- mirrors src/pipeline.py, see module docstring
# ---------------------------------------------------------------------------

class Resolution:
    """Outcome of re-classifying one entry."""

    def __init__(self, kind, **fields):
        self.kind = kind  # "filed" | "junk" | "unresolved"
        self.__dict__.update(fields)


def classify_entry(email):
    """Runs the same decision cascade as src/pipeline.py's per-email loop
    (minus anything that needs Outlook/live To-Cc data -- see the module
    docstring) and returns a Resolution describing what should happen."""
    if is_internal_sender(email) or is_ignored_sender(email):
        return Resolution("junk", reason="internal or ignored sender")

    email_year = email["timestamp"].year

    # Deterministic US-code match, tried regardless of sender -- free,
    # unambiguous string match, no AI call.
    us_codes = extract_us_codes(email)
    for code in us_codes:
        resolved = resolve_plenergy_folder(ARCHIVE_ROOT, email_year, code)
        if resolved:
            project_folder_name, address_folder_name = resolved
            sender_lower = (email.get("sender") or "").lower()
            contact_label = "PLAINCO" if "plainco.es" in sender_lower else "PLENERGY"
            return Resolution(
                "filed",
                project_folder_name=project_folder_name,
                address_folder_name=address_folder_name,
                contact_label=contact_label,
                topic_label=code,
                company_only=False,
                existing_company=True,
                summary=cheap_fallback_summary(email),
            )

    if is_plenergy_sender(email):
        contact_label = "PLAINCO" if "plainco.es" in (email.get("sender") or "").lower() else "PLENERGY"
        do_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, DO_PLENERGY_FOLDER)
        project_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, PLENERGY_FOLDER)
        try:
            llm_match = classify_plenergy_address(email, do_candidates, project_candidates)
        except Exception as e:
            return Resolution("unresolved", reason=f"Plenergy classification failed: {e}")

        if llm_match.matched_existing:
            topic_label = us_codes[0] if len(us_codes) == 1 else "ESTACION IDENTIFICADA"
            return Resolution(
                "filed",
                project_folder_name=llm_match.matched_folder,
                address_folder_name=llm_match.address_folder_name,
                contact_label=contact_label,
                topic_label=topic_label,
                company_only=False,
                existing_company=True,
                summary=llm_match.summary,
            )

        company_candidates = list_existing_companies(ARCHIVE_ROOT, email_year)
        plenergy_company = exact_existing_name(PLENERGY_FOLDER, company_candidates)
        topic_for_index = (
            (llm_match.address_folder_name.strip() if llm_match.address_folder_name else "")
            or (us_codes[0] if us_codes else "PROYECTO SIN IDENTIFICAR")
        )
        if plenergy_company:
            return Resolution(
                "filed",
                project_folder_name=plenergy_company,
                address_folder_name=None,
                contact_label=contact_label,
                topic_label=topic_for_index,
                company_only=True,
                existing_company=True,
                summary=llm_match.summary,
            )
        return Resolution(
            "filed",
            project_folder_name="PLENERGY",
            address_folder_name=None,
            contact_label=contact_label,
            topic_label=topic_for_index,
            company_only=False,
            existing_company=False,
            summary=llm_match.summary,
        )

    # General company-first routing.
    existing_companies = list_existing_companies(ARCHIVE_ROOT, email_year)
    try:
        company_match = classify_project(email, existing_companies)
    except Exception as e:
        return Resolution("unresolved", reason=f"Company classification failed: {e}")

    if not company_match.is_relevant:
        return Resolution("junk", reason="Gemini judged it irrelevant (spam/marketing/automated/personal)")

    contact_label = company_match.contact_label
    topic_label = company_match.topic_label
    summary = company_match.summary
    real_company_name = exact_existing_name(company_match.project_folder_name, existing_companies)

    address_folder_name = None
    company_only = False
    existing_company = False
    project_folder_name = None

    if real_company_name:
        existing_company = True
        project_folder_name = real_company_name
        company_year = get_project_year(project_folder_name) or email_year
        existing_projects = list_existing_addresses(ARCHIVE_ROOT, company_year, project_folder_name)

        if existing_projects:
            try:
                project_match = classify_address(email, existing_projects, project_folder_name)
                real_project_name = (
                    exact_existing_name(project_match.address_folder_name, existing_projects)
                    if project_match.matched_existing else None
                )
            except Exception as e:
                return Resolution("unresolved", reason=f"Project/site classification failed: {e}")

            if real_project_name:
                address_folder_name = real_project_name
                company_only = False
            else:
                company_only = True
        else:
            company_only = True
    else:
        project_folder_name = (
            str(company_match.project_folder_name or "").strip()
            or str(company_match.contact_label or "").strip()
            or "DESCONOCIDO"
        )

    # Plenergy/DO PLENERGY rescue check, same as pipeline.py.
    if (not is_plenergy_sender(email)) and ((not existing_company) or company_only):
        try:
            do_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, DO_PLENERGY_FOLDER)
            project_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, PLENERGY_FOLDER)
            rescue_match = classify_plenergy_address(email, do_candidates, project_candidates)
        except Exception:
            rescue_match = None

        if rescue_match is not None and rescue_match.matched_existing:
            project_folder_name = rescue_match.matched_folder
            address_folder_name = rescue_match.address_folder_name
            topic_label = us_codes[0] if len(us_codes) == 1 else "ESTACION IDENTIFICADA"
            summary = rescue_match.summary or summary
            existing_company = True
            company_only = False

    return Resolution(
        "filed",
        project_folder_name=project_folder_name,
        address_folder_name=address_folder_name,
        contact_label=contact_label,
        topic_label=topic_label,
        company_only=company_only,
        existing_company=existing_company,
        summary=summary,
    )


# ---------------------------------------------------------------------------
# Filesystem actions
# ---------------------------------------------------------------------------

def _move_unique(src_dir: str, dest_parent_dir: str, folder_name: str, dry_run: bool) -> str:
    """Picks a free destination folder name (matching the same " (2)",
    " (3)" collision convention as save_email._make_unique_folder) and,
    unless dry_run, actually moves src_dir there."""
    candidate = os.path.join(dest_parent_dir, folder_name)
    counter = 2
    while os.path.isdir(candidate):
        candidate = os.path.join(dest_parent_dir, f"{folder_name} ({counter})")
        counter += 1
    if not dry_run:
        os.makedirs(dest_parent_dir, exist_ok=True)
        shutil.move(src_dir, candidate)
    return candidate


def _count_attachment_files(entry_path: str) -> int:
    skip = {"email.txt", "email.msg", "email.pdf", "metadata.json"}
    try:
        return sum(1 for name in os.listdir(entry_path) if name not in skip and os.path.isfile(os.path.join(entry_path, name)))
    except OSError:
        return 0


# ---------------------------------------------------------------------------
# CSV bookkeeping -- update-in-place by EntryID, never duplicate a row
# ---------------------------------------------------------------------------

def _read_csv_rows(path):
    if not os.path.isfile(path):
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _write_csv_rows(path, fieldnames, rows):
    fd, temp_name = tempfile.mkstemp(prefix="reprocess_", suffix=".csv", dir=os.path.dirname(path) or ".")
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with temp_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in rows:
                writer.writerow({name: row.get(name, "") for name in fieldnames})
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


def upsert_index_row(output_root, entry_id, row_dict, dry_run):
    path = os.path.join(output_root, "index.csv")
    rows = _read_csv_rows(path)
    updated = False
    if entry_id:
        for row in rows:
            if str(row.get("EntryID") or "").strip() == entry_id:
                row.update(row_dict)
                updated = True
                break
    if not updated:
        rows.append(row_dict)
    if not dry_run:
        _write_csv_rows(path, INDEX_FIELDNAMES, rows)
    return updated


def upsert_report_row(output_root, entry_id, row_dict, dry_run):
    path = os.path.join(output_root, "email_report_log.csv")
    rows = _read_csv_rows(path)
    updated = False
    if entry_id:
        for row in rows:
            if str(row.get("EntryID") or "").strip() == entry_id:
                row.update(row_dict)
                updated = True
                break
    if not updated:
        rows.append(row_dict)
    if not dry_run:
        _write_csv_rows(path, REPORT_FIELDNAMES, rows)
    return updated


def remove_rows_by_entry_id(path, fieldnames, entry_id, dry_run):
    if not entry_id or not os.path.isfile(path):
        return
    rows = _read_csv_rows(path)
    kept = [row for row in rows if str(row.get("EntryID") or "").strip() != entry_id]
    if len(kept) != len(rows) and not dry_run:
        _write_csv_rows(path, fieldnames, kept)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(dry_run: bool = False):
    ensure_output_root()
    year = datetime.datetime.now().year
    unsorted_root = _unsorted_root(year)
    if not os.path.isdir(unsorted_root):
        print(f"No UNSORTED folder found for {year}: {unsorted_root}")
        return

    entries = list(_find_entries(unsorted_root))
    print(f"Found {len(entries)} entr{'y' if len(entries) == 1 else 'ies'} in {unsorted_root}"
          + (" (DRY RUN -- nothing will be moved or deleted)" if dry_run else ""))

    moved = deleted = skipped = unreadable = 0

    for entry_path in entries:
        email = load_entry(entry_path)
        if email is None:
            print(f"  SKIP (couldn't read email.txt/metadata.json): {entry_path}")
            unreadable += 1
            continue

        if email["direction"] == "SALIENTE":
            # Outgoing mail is never processed or shown anywhere anymore --
            # leave any legacy SALIENTE entry exactly where it is.
            continue

        entry_id = email["id"]
        label = email["subject"]

        try:
            result = classify_entry(email)
        except Exception as e:
            print(f"  UNRESOLVED (unexpected error, left in place): {label} -- {e}")
            skipped += 1
            continue

        if result.kind == "unresolved":
            print(f"  UNRESOLVED (still can't classify -- {result.reason}): {label}")
            skipped += 1
            continue

        if result.kind == "junk":
            print(f"  DELETE ({result.reason}): {label}")
            if not dry_run:
                shutil.rmtree(entry_path, ignore_errors=True)
                remove_rows_by_entry_id(os.path.join(OUTPUT_ROOT, "index.csv"), INDEX_FIELDNAMES, entry_id, dry_run)
                remove_rows_by_entry_id(os.path.join(OUTPUT_ROOT, "email_report_log.csv"), REPORT_FIELDNAMES, entry_id, dry_run)
                remove_rows_by_entry_id(os.path.join(OUTPUT_ROOT, "priorities.csv"), ["EntryID", "Priority", "Date", "Subject"], entry_id, dry_run)
                remove_rows_by_entry_id(os.path.join(OUTPUT_ROOT, "pendientes.csv"), ["EntryID", "Date", "Sender", "Subject"], entry_id, dry_run)
            deleted += 1
            continue

        # result.kind == "filed"
        dest_path = resolve_save_destination(
            email, result.project_folder_name, result.contact_label, ARCHIVE_ROOT,
            result.address_folder_name, result.company_only, result.existing_company,
        )
        dest_parent = os.path.dirname(dest_path)
        folder_name = os.path.basename(dest_path)
        final_path = _move_unique(entry_path, dest_parent, folder_name, dry_run)
        print(f"  MOVE -> {final_path}\n         ({label})")

        if not dry_run:
            meta_path = os.path.join(final_path, "metadata.json")
            if os.path.isfile(meta_path):
                try:
                    with open(meta_path, encoding="utf-8") as f:
                        meta = json.load(f)
                except (OSError, json.JSONDecodeError):
                    meta = {}
                meta.update({
                    "project_folder": result.project_folder_name,
                    "address_folder": result.address_folder_name,
                    "company_only": bool(result.company_only),
                    "existing_company": bool(result.existing_company),
                    "contact_label": result.contact_label,
                    "topic_label": result.topic_label,
                })
                with open(meta_path, "w", encoding="utf-8") as f:
                    json.dump(meta, f, indent=2, ensure_ascii=False)

            upsert_index_row(OUTPUT_ROOT, entry_id, {
                "EntryID": entry_id,
                "StoreID": email.get("store_id") or "",
                "Year": email["timestamp"].year,
                "Project Folder": result.project_folder_name,
                "Address Folder": result.address_folder_name or "",
                "Direction": email["direction"],
                "Contact": result.contact_label,
                "Sender Name": email.get("sender_name") or "",
                "Sender Email": email.get("sender_email") or email.get("sender") or "",
                "Topic": result.topic_label,
                "Date": email["timestamp"].strftime("%Y-%m-%d %H:%M"),
                "Subject": email["subject"],
                "Sender/Recipient": email.get("sender") or email.get("recipient") or "",
                "Attachments": _count_attachment_files(final_path),
                "Folder Path": final_path,
            }, dry_run)

            upsert_report_row(OUTPUT_ROOT, entry_id, {
                "EntryID": entry_id,
                "Date": email["timestamp"].strftime("%Y-%m-%d"),
                "Time": email["timestamp"].strftime("%H:%M"),
                "Sender Name": email.get("sender_name") or "",
                "Sender Email": email.get("sender_email") or email.get("sender") or email.get("recipient") or "",
                "Subject": email["subject"],
                "Summary": result.summary or "",
                "Path": final_path,
            }, dry_run)

        moved += 1

    print(
        f"\nDone. {moved} moved, {deleted} deleted, {skipped} still unresolved, "
        f"{unreadable} unreadable, out of {len(entries)} entries checked."
    )

    if not dry_run and (moved or deleted):
        generate_status_page(OUTPUT_ROOT)
        generate_email_report_xlsx(OUTPUT_ROOT)
        print("Status page and Informe de Emails.xlsx regenerated.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="Preview the plan without moving or deleting anything.")
    args = parser.parse_args()
    run(dry_run=args.dry_run)
