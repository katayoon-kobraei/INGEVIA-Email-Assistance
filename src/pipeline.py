from src.config import OUTPUT_ROOT, ensure_output_root
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.project_folders import list_existing_projects, get_next_project_code
from src.classification.project_agent import classify_project
from src.output.index_writer import append_to_index


RELEVANT_YEARS = [2026, 2025]

def run():
    ensure_output_root()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
            year = email["timestamp"].year
            existing = list_existing_projects(OUTPUT_ROOT, RELEVANT_YEARS)
            try:
                match = classify_project(email, existing)
                if match.matched_existing:
                    project_folder_name = match.project_folder_name
                else:
                    code = get_next_project_code(OUTPUT_ROOT, year)
                    project_folder_name = f"{code} {match.project_folder_name}"
                contact_label = match.contact_label
                topic_label = match.topic_label
            except Exception as e:
                print(f"Project classification failed for {email['subject']}: {e}")
                project_folder_name, contact_label, topic_label = "UNSORTED", "DESCONOCIDO", "SIN CLASIFICAR"
                
            folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT)
            append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT)
            mark_processed(email["id"], OUTPUT_ROOT)
            print(f"Saved: {email['subject']} -> {folder}")
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

if __name__ == "__main__":
    run()