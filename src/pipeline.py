from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
)
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import archive_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.archive_state import load_pending_archive, add_pending_archive, remove_pending_archive
from src.output.pending_list import append_to_pending_list
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.pending_agent import classify_pending
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.index_writer import append_to_index
from src.output.status_page import generate_status_page

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def _retry_pending_flags():
    """Emails whose Outlook flag failed on a previous run (usually
    because that exact email was open/selected in Outlook at the
    time) get retried here, before anything new is processed."""
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
    """Retries pending-response emails whose move to PENDING_FOLDER_NAME
    failed last run."""
    if not CHECK_PENDING_RESPONSES:
        return
    pending = load_pending_archive(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose folder move failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if archive_email(entry_id, folder_name):
            remove_pending_archive(OUTPUT_ROOT, entry_id)
            print(f"  Moved on retry: {entry_id} -> {folder_name}")


def _mark_done(email):
    """Marks an email as processed locally (dedupe state) and, if
    enabled, stamps it back in Outlook so the boss can see it was
    handled. If the Outlook write fails (e.g. the email was
    open/selected at that moment), it's queued to retry next run."""
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        return
    ok = mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
    else:
        add_pending_flag(OUTPUT_ROOT, email["id"], PROCESSED_CATEGORY_NAME)
        print(f"  (Outlook flag failed -- will retry automatically next run)")


def _handle_pending_check(email):
    """After a relevant email is filed, run a cheap separate check for
    whether it's still waiting on a written reply. Only meaningful for
    incoming mail. If it needs a reply: log it to pendientes.csv and
    move it (in Outlook) into PENDING_FOLDER_NAME."""
    if not CHECK_PENDING_RESPONSES or email.get("direction") != "ENTRANTE":
        return
    try:
        pending = classify_pending(email)
    except Exception as e:
        print(f"Pending-response check failed for {email['subject']}: {e}")
        return
    if not pending.needs_response:
        return

    append_to_pending_list(email, OUTPUT_ROOT)
    ok = archive_email(email["id"], PENDING_FOLDER_NAME)
    if ok:
        remove_pending_archive(OUTPUT_ROOT, email["id"])
    else:
        add_pending_archive(OUTPUT_ROOT, email["id"], PENDING_FOLDER_NAME)
        print(f"  (Move to pending folder failed -- will retry automatically next run)")
    print(f"  Marked as PENDING RESPONSE: {email['subject']}")


def run():
    ensure_output_root()
    _retry_pending_flags()
    _retry_pending_archives()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
            email_year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            try:
                # classify_project now decides relevance itself -- no
                # separate pre-filter call, no separate token cost for
                # it. If it's not worth filing (junk, marketing, social
                # notifications, etc.), leave it completely untouched
                # in the Inbox: no flag, no move, no save. Staff
                # handle that manually. Still marked processed so it's
                # classified exactly once, never re-checked (and
                # re-billed) on future runs.
                match = classify_project(email, existing)

                if not match.is_relevant:
                    mark_processed(email["id"], OUTPUT_ROOT)
                    print(f"Ignored (not relevant): {email['subject']}")
                    continue

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
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)


if __name__ == "__main__":
    run()