from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email

import win32com.client
import datetime


if __name__ == "__main__":
    for email in get_recent_emails(60):
        folder = save_email(email, r"C:\EmailAssistant\Output")
        print("Saved to:", folder)