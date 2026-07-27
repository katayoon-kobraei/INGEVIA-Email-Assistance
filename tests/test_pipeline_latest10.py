import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import win32com.client

from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
    BOSS_EMAIL,
    ADMINISTRACION_EMAIL,
    BILLING_OUTPUT_ROOT,
)
from src.output.save_email import save_email, save_billing_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import copy_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.pending_copy_state import load_pending_copies, add_pending_copy, remove_pending_copy
from src.output.pending_list import append_to_pending_list
from src.output.priority_list import append_to_priority_list
from src.output.billing_routing import is_external_sender, boss_is_recipient, administracion_is_recipient
from src.output.outlook_forward import forward_email
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.billing_agent import classify_billing
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
                "to": message.To,
                "cc": message.CC,
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
    """Cheap urgency score (1-5) for anything that got filed. Purely
    informational -- doesn't affect filing or Outlook state, just logs
    to priorities.csv for the UI to read."""
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
    _retry_pending_copies()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = _get_latest_inbox_emails(COUNT)
    print(f"Testing pipeline on the {len(emails)} latest Inbox email(s) (ignoring time), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            print(f"Skipping (already processed): {email['subject']}")
            continue
        try:
            # Billing/procurement check -- same gate as src/pipeline.py:
            # external sender, boss on To/Cc. Handled separately from
            # normal project filing, and skips the rest of the loop.
            if is_external_sender(email) and boss_is_recipient(email):
                try:
                    billing = classify_billing(email)
                except Exception as e:
                    print(f"Billing check failed for {email['subject']}: {e}")
                    billing = None

                if billing is not None and billing.is_billing_related:
                    folder = save_billing_email(email, BILLING_OUTPUT_ROOT)
                    mark_processed(email["id"], OUTPUT_ROOT)
                    if administracion_is_recipient(email):
                        print(f"Saved (billing, admin already on it): {email['subject']} -> {folder}")
                    else:
                        ok = forward_email(email["id"], ADMINISTRACION_EMAIL)
                        if ok:
                            print(f"Saved (billing) and forwarded to {ADMINISTRACION_EMAIL}: {email['subject']} -> {folder}")
                        else:
                            print(f"Saved (billing) but forwarding FAILED: {email['subject']} -> {folder}")
                    continue

            email_year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            try:
                # classify_project decides relevance itself now -- if
                # it's not worth filing, leave it completely untouched
                # in the Inbox (no flag, no move, no save).
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
            _handle_priority_check(email)
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)
    print("Done.")


if __name__ == "__main__":
    run()