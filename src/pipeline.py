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
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import archive_email, copy_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.archive_state import load_pending_archive, add_pending_archive, remove_pending_archive
from src.output.pending_copy_state import load_pending_copies, add_pending_copy, remove_pending_copy
from src.output.pending_list import append_to_pending_list
from src.output.department_routing import match_department
from src.output.outlook_archive import archive_email, copy_email, archive_to_top_level
from src.output.save_email import save_email, save_department_email
from src.classification.priority_agent import classify_priority
from src.output.priority_list import append_to_priority_list
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.relevance_agent import classify_relevance
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
    """Same idea as _retry_pending_flags(), but for emails whose move
    to another Outlook folder failed last run -- this is junk mail
    only now (moved to JUNK_ARCHIVE_FOLDER_NAME); pending-response
    emails use their own copy-retry queue below instead."""
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
    """Same idea, but for pending-response emails whose COPY into
    PENDING_FOLDER_NAME failed last run (e.g. Outlook was busy on
    that item at that moment). The original stays in the Inbox either
    way -- this only retries getting the copy into the pending folder."""
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
    """Marks an email as processed locally (dedupe state) and, if
    enabled, stamps it back in Outlook so the boss can see it was
    handled. Used for emails that got filed into a project. If the
    Outlook write fails (e.g. the email was open/selected at that
    moment), it's queued to retry automatically next run."""
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        return
    ok = mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
    else:
        add_pending_flag(OUTPUT_ROOT, email["id"], PROCESSED_CATEGORY_NAME)
        print(f"  (Outlook flag failed -- will retry automatically next run)")


def _handle_junk(email):
    mark_processed(email["id"], OUTPUT_ROOT)

    if not ARCHIVE_JUNK_EMAILS:
        if FLAG_PROCESSED_EMAILS:
            mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
        return

    ok = archive_to_top_level(email["id"], JUNK_ARCHIVE_FOLDER_NAME)
    if ok:
        remove_pending_archive(OUTPUT_ROOT, email["id"])
    else:
        add_pending_archive(OUTPUT_ROOT, email["id"], JUNK_ARCHIVE_FOLDER_NAME)
        print(f"  (Archive move failed -- will retry automatically next run)")

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


def _handle_pending_check(email):
    """After a relevant email is filed, run a cheap separate check for
    whether it's still waiting on a written reply. Only meaningful for
    incoming mail -- something the firm itself sent doesn't need a
    reply FROM the firm. If it needs a reply: log it to pendientes.csv
    and put a COPY of it (in Outlook) into PENDING_FOLDER_NAME -- the
    original stays in the Inbox, unlike junk archiving which moves the
    original out."""
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
    ok = copy_email(email["id"], PENDING_FOLDER_NAME)
    if ok:
        remove_pending_copy(OUTPUT_ROOT, email["id"])
    else:
        add_pending_copy(OUTPUT_ROOT, email["id"], PENDING_FOLDER_NAME)
        print(f"  (Copy to pending folder failed -- will retry automatically next run)")
    print(f"  Marked as PENDING RESPONSE: {email['subject']}")


def run():
    ensure_output_root()
    _retry_pending_flags()
    _retry_pending_archives()
    _retry_pending_copies()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
            # Deterministic department routing (no LLM, no cost) --
            # runs before anything else. If the sender's domain
            # matches a configured department (e.g. the secretary),
            # file it straight into that department's folder and skip
            # relevance/project/address/pending classification
            # entirely -- this isn't client correspondence.
            department_folder_name = match_department(email)
            if department_folder_name:
                folder = save_department_email(email, department_folder_name, OUTPUT_ROOT)
                _mark_done(email)
                ok = copy_email(email["id"], department_folder_name)
                if ok:
                    remove_pending_copy(OUTPUT_ROOT, email["id"])
                else:
                    add_pending_copy(OUTPUT_ROOT, email["id"], department_folder_name)
                    print(f"  (Copy to '{department_folder_name}' failed -- will retry automatically next run)")
                print(f"Saved (department: {department_folder_name}): {email['subject']} -> {folder}")
                continue

            # Cheap first-pass filter: a short, separate prompt with
            # no folder structure, no matching rules -- just "is this
            # real correspondence or noise?". This runs on EVERY
            # email so junk (social media notifications, marketing,
            # automated mail) gets caught for a fraction of the token
            # cost of the full classify_project prompt below, instead
            # of paying for that long prompt on every single email.
            # If the relevance check itself fails (API error), we
            # fall through to full classification rather than risk
            # silently dropping a real email.
            try:
                relevance = classify_relevance(email)
            except Exception as e:
                print(f"Relevance check failed for {email['subject']}: {e}")
                relevance = None

            if relevance is not None and not relevance.is_relevant:
                _handle_junk(email)
                print(f"Skipped (not relevant): {email['subject']}")
                continue

            # Matching is scoped to the email's own year only -- a
            # returning client's folder from a different year will not
            # be shown as a candidate here.
            email_year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            try:
                # classify_project just returns a name -- if it's a
                # real match, it already has a code; if not, it's a
                # bare name and save_email() -> resolve_project_relative_path()
                # routes it into that year's holding pen and mints its
                # local "{NNN}" counter automatically, without minting
                # a new formal project code on its own.
                match = classify_project(email, existing)
                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label

                # Address sub-folders apply to formal AND holding-pen
                # companies alike -- only the CODE differs: a formal
                # company gets a real "{code}-{NN}" address code, a
                # holding-pen company just gets a bare address name
                # (no code sequence exists yet for it).
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