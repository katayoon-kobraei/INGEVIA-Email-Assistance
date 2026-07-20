from src.config import OUTPUT_ROOT, ensure_output_root
from src.ingestion.outlook_local import get_recent_emails
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


def run():
    ensure_output_root()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
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

                # Noise filter: social media notifications, marketing,
                # automated system mail, etc. -- not real client
                # correspondence, so it's skipped entirely (not even
                # sent to UNSORTED/holding pen), but still marked
                # processed so it isn't re-evaluated every run.
                if not match.is_relevant:
                    mark_processed(email["id"], OUTPUT_ROOT)
                    print(f"Skipped (not relevant): {email['subject']}")
                    continue

                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label

                # Second step: if this email names a specific site,
                # match it against that company's existing sites (or
                # mint/propose the next one automatically -- safe to do
                # without a human gate, since the company itself was
                # already vetted). Works both for a company that
                # already has a real code (numeric "{code}-{NN}"
                # addresses) and for a not-yet-formal holding-pen
                # company (bare address names, no code yet).
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
                            # Not-yet-formal company: no code sequence
                            # to mint from yet, so just use the bare
                            # address name directly.
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


if __name__ == "__main__":
    run()