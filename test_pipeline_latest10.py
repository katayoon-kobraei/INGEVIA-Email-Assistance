import win32com.client

from src.config import OUTPUT_ROOT, ensure_output_root
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
)
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.index_writer import append_to_index
from src.output.status_page import generate_status_page

INBOX_FOLDER_ID = 6
COUNT = 10


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


def run():
    ensure_output_root()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = _get_latest_inbox_emails(COUNT)
    print(f"Testing pipeline on the {len(emails)} latest Inbox email(s) (ignoring time), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            print(f"Skipping (already processed): {email['subject']}")
            continue
        try:
            email_year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            try:
                match = classify_project(email, existing)

                if not match.is_relevant:
                    mark_processed(email["id"], OUTPUT_ROOT)
                    print(f"Skipped (not relevant): {email['subject']}")
                    continue

                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label

                if match.mentions_specific_address:
                    company_year = get_project_year(project_folder_name) or email_year
                    existing_addresses = list_existing_addresses(OUTPUT_ROOT, company_year, project_folder_name)
                    try:
                        addr_match = classify_address(email, existing_addresses)
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
            mark_processed(email["id"], OUTPUT_ROOT)
            print(f"Saved: {email['subject']} -> {folder}")
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)
    print("Done.")


if __name__ == "__main__":
    run()