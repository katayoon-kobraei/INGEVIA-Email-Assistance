from src.config import OUTPUT_ROOT, ensure_output_root
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.classification.department_agent import classify_department

def run():
    ensure_output_root()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
            try:
                department = classify_department(email).department.value
            except Exception as e:
                print(f"Classification failed for {email['subject']}: {e}")
                department = "Unclassified"

            folder = save_email(email, department, OUTPUT_ROOT)
            mark_processed(email["id"], OUTPUT_ROOT)
            print(f"Saved: {email['subject']} -> {folder}")
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

if __name__ == "__main__":
    run()