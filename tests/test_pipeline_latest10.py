import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import win32com.client

from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    ARCHIVE_JUNK_EMAILS,
    JUNK_ARCHIVE_FOLDER_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
)
from src.output.save_email import save_email, save_department_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import archive_email, copy_email, archive_to_top_level
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.archive_state import load_pending_archive, add_pending_archive, remove_pending_archive
from src.output.pending_list import append_to_pending_list
from src.output.pending_copy_state import load_pending_copies, add_pending_copy, remove_pending_copy
from src.output.priority_list import append_to_priority_list
from src.output.department_routing import match_department
from src.output.department_routing import match_department, should_ignore
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.relevance_agent import classify_relevance
from src.classification.pending_agent import classify_pending
from src.classification.priority_agent import classify_priority
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.index_writer import append_to_index
from src.output.status_page import generate_status_page

INBOX_FOLDER_ID = 6
COUNT = 3


def _get_latest_inbox_emails(count):
    """Local to this script only -- does not touch outlook_local.py."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    results = []
    for message in messages:
        if len(results) >= count:
            break
        try:
            results.append({
                "id": message.EntryID,
                "subject": message.Subject,
                "sender": message.SenderEmailAddress,
                "recipient": None,
                "timestamp": message.ReceivedTime,
                "body": message.Body,
                "attachments": message.Attachments,
                "direction": "ENTRANTE",
            })
        except Exception:
            continue  # skip odd item types (meeting requests, read receipts, etc.)
    return results


def _retry_pending_flags():
    if not FLAG_PROCESSED_EMAILS:
        return
    pending = load_pending_flags(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose Outlook flag failed last run...")
    for entry_id, category in list(pending.items()):
        if mark_email_processed(entry_id, category):
            remove_pending_flag(OUTPUT_ROOT, entry_id)
            print(f"  Flagged on retry: {entry_id}")


def _retry_pending_archives():
    if not ARCHIVE_JUNK_EMAILS:
        return
    pending = load_pending_archive(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose folder move failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if archive_email(entry_id, folder_name):
            remove_pending_archive(OUTPUT_ROOT, entry_id)
            print(f"  Moved on retry: {entry_id} -> {folder_name}")


def _retry_pending_copies():
    if not CHECK_PENDING_RESPONSES:
        return
    pending = load_pending_copies(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} pending-email copy(ies) that failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if copy_email(entry_id, folder_name):
            remove_pending_copy(OUTPUT_ROOT, entry_id)
            print(f"  Copied on retry: {entry_id} -> {folder_name}")


def _mark_done(email):
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        print("  (Outlook flag skipped -- FLAG_PROCESSED_EMAILS is off in .env)")
        return
    ok = mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
        print("  (Outlook flag: applied OK)")
    else:
        add_pending_flag(OUTPUT_ROOT, email["id"], PROCESSED_CATEGORY_NAME)
        print("  (Outlook flag: FAILED -- queued to retry next run)")


def _handle_junk(email):
    mark_processed(email["id"], OUTPUT_ROOT)

    if not ARCHIVE_JUNK_EMAILS:
        print("  (Archive skipped -- ARCHIVE_JUNK_EMAILS is off in .env)")
        if FLAG_PROCESSED_EMAILS:
            mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
        return

    ok = archive_to_top_level(email["id"], JUNK_ARCHIVE_FOLDER_NAME)
    if ok:
        remove_pending_archive(OUTPUT_ROOT, email["id"])
        print(f"  (Moved to '{JUNK_ARCHIVE_FOLDER_NAME}' OK)")
    else:
        add_pending_archive(OUTPUT_ROOT, email["id"], JUNK_ARCHIVE_FOLDER_NAME)
        print("  (Archive move FAILED -- queued to retry next run)")


def _handle_pending_check(email):
    if not CHECK_PENDING_RESPONSES or email.get("direction") != "ENTRANTE":
        return
    try:
        pending = classify_pending(email)
    except Exception as e:
        print(f"Pending-response check failed for {email['subject']}: {e}")
        return
    if not pending.needs_response:
        print("  (Not pending -- no reply needed)")
        return

    append_to_pending_list(email, OUTPUT_ROOT)
    ok = copy_email(email["id"], PENDING_FOLDER_NAME)
    if ok:
        remove_pending_copy(OUTPUT_ROOT, email["id"])
        print(f"  (Copied to '{PENDING_FOLDER_NAME}' OK -- logged to pendientes.csv)")
    else:
        add_pending_copy(OUTPUT_ROOT, email["id"], PENDING_FOLDER_NAME)
        print("  (Copy to pending folder FAILED -- queued to retry next run)")


def _handle_priority_check(email):
    """Cheap urgency score (1-5) for anything that passed the junk
    filter. Purely informational -- doesn't affect filing or Outlook
    state, just logs to priorities.csv for the UI to read."""
    try:
        result = classify_priority(email)
    except Exception as e:
        print(f"Priority check failed for {email['subject']}: {e}")
        return
    append_to_priority_list(email, result.priority, OUTPUT_ROOT)
    print(f"  Priority: {result.priority}/5")


def run():
    ensure_output_root()
    _retry_pending_flags()
    _retry_pending_archives()
    _retry_pending_copies()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = _get_latest_inbox_emails(COUNT)
    print(f"Testing pipeline on the {len(emails)} latest Inbox email(s) (ignoring time), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            print(f"Skipping (already processed): {email['subject']}")
            continue
        try:
            department_folder_name = match_department(email)
            if department_folder_name:
                folder = save_department_email(email, department_folder_name, OUTPUT_ROOT)
                _mark_done(email)
                ok = copy_email(email["id"], department_folder_name)
                if ok:
                    remove_pending_copy(OUTPUT_ROOT, email["id"])
                    print(f"  (Copied to '{department_folder_name}' OK)")
                else:
                    add_pending_copy(OUTPUT_ROOT, email["id"], department_folder_name)
                    print(f"  (Copy to '{department_folder_name}' failed -- queued to retry next run)")
                print(f"Saved (department: {department_folder_name}): {email['subject']} -> {folder}")
                continue

            if should_ignore(email):
                mark_processed(email["id"], OUTPUT_ROOT)
                print(f"Ignored (internal domain): {email['subject']}")
                continue

            try:
                relevance = classify_relevance(email)
            except Exception as e:
                print(f"Relevance check failed for {email['subject']}: {e}")
                relevance = None

            if relevance is not None and not relevance.is_relevant:
                _handle_junk(email)
                print(f"Skipped (not relevant): {email['subject']}")
                continue

            email_year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            try:
                match = classify_project(email, existing)
                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label

                if match.mentions_specific_address:
                    company_year = get_project_year(project_folder_name) or email_year
                    existing_addresses = list_existing_addresses(OUTPUT_ROOT, company_year, project_folder_name)
                    try:
                        addr_match = classify_address(email, existing_addresses, project_folder_name)
                        if addr_match.matched_existing:
                            address_folder_name = addr_match.address_folder_name
                        elif is_formal_project_code(project_folder_name):
                            addr_code = get_next_address_code(OUTPUT_ROOT, company_year, project_folder_name)
                            address_folder_name = f"{addr_code} {addr_match.address_folder_name}"
                        else:
                            address_folder_name = addr_match.address_folder_name
                    except Exception as e:
                        print(f"Address classification failed for {email['subject']}: {e}")
                        address_folder_name = None
            except Exception as e:
                print(f"Project classification failed for {email['subject']}: {e}")
                project_folder_name, contact_label, topic_label = "UNSORTED", "DESCONOCIDO", "SIN CLASIFICAR"

            folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT, address_folder_name)
            append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT, address_folder_name)
            _mark_done(email)
            print(f"Saved: {email['subject']} -> {folder}")

            _handle_pending_check(email)
            _handle_priority_check(email)
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)
    print("Done.")


if __name__ == "__main__":
    run()