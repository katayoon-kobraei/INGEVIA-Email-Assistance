from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    ARCHIVE_JUNK_EMAILS,
    JUNK_ARCHIVE_FOLDER_NAME,
)
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import archive_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.archive_state import load_pending_archive, add_pending_archive, remove_pending_archive
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.relevance_agent import classify_relevance
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.index_writer import append_to_index
from src.output.status_page import generate_status_page


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
    """Same idea as _retry_pending_flags(), but for junk emails whose
    move-to-archive-folder failed last run."""
    if not ARCHIVE_JUNK_EMAILS:
        return
    pending = load_pending_archive(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose archive move failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if archive_email(entry_id, folder_name):
            remove_pending_archive(OUTPUT_ROOT, entry_id)
            print(f"  Archived on retry: {entry_id}")


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
    """Marks a junk/irrelevant email as processed locally, then either
    moves it out of the Inbox into the archive subfolder (default) or
    just flags it in place if ARCHIVE_JUNK_EMAILS is off. A failed
    move is queued to retry automatically next run, same as flags."""
    mark_processed(email["id"], OUTPUT_ROOT)

    if not ARCHIVE_JUNK_EMAILS:
        if FLAG_PROCESSED_EMAILS:
            mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
        return

    ok = archive_email(email["id"], JUNK_ARCHIVE_FOLDER_NAME)
    if ok:
        remove_pending_archive(OUTPUT_ROOT, email["id"])
    else:
        add_pending_archive(OUTPUT_ROOT, email["id"], JUNK_ARCHIVE_FOLDER_NAME)
        print(f"  (Archive move failed -- will retry automatically next run)")


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
                # No more automatic new-project-code creation here.
                # classify_project just returns a name -- if it's a
                # real match, it already has a code; if not, it's a
                # bare name and save_email() will route it into that
                # year's holding pen ("{yy}-000 MAILS") instead of
                # minting a new formal project on its own.
                match = classify_project(email, existing)
                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label

                # Second step, only for a company that already has a
                # real code: if this email names a specific site, match
                # it against that company's existing sites (or mint the
                # next one automatically -- safe to do without a human
                # gate, since the company itself was already vetted).
                if match.mentions_specific_address and is_formal_project_code(project_folder_name):
                    company_year = get_project_year(project_folder_name) or email_year
                    existing_addresses = list_existing_addresses(OUTPUT_ROOT, company_year, project_folder_name)
                    try:
                        addr_match = classify_address(email, existing_addresses)
                        if addr_match.matched_existing:
                            address_folder_name = addr_match.address_folder_name
                        else:
                            addr_code = get_next_address_code(OUTPUT_ROOT, company_year, project_folder_name)
                            address_folder_name = f"{addr_code} {addr_match.address_folder_name}"
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
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)


if __name__ == "__main__":
    run()