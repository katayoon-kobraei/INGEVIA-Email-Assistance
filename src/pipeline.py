from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed

OUTPUT_ROOT = r"U:\Ingevia\INGEVIA-Email-Assistanc\storage"

def run():
    processed = load_processed_ids(OUTPUT_ROOT)
    for email in get_recent_emails(60):
        if email["id"] in processed:
            continue
        try:
            save_email(email, OUTPUT_ROOT)
            mark_processed(email["id"], OUTPUT_ROOT)
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue  # don't mark it — next run will retry this one specifically

if __name__ == "__main__":
    run()