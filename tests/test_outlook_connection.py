from tests.test_outlook_connection import get_recent_emails  # or wherever you moved it
from src.output.save_email import save_email

import win32com.client
import datetime

def get_recent_emails(minutes_back=30):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)          # 6 = the Inbox folder
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)         # newest first

    cutoff = datetime.datetime.now() - datetime.timedelta(minutes=minutes_back)
    results = []

    for message in messages:
        received = message.ReceivedTime.replace(tzinfo=None)
        if received < cutoff:
            break  # sorted newest-first, so nothing after this matters

        results.append({
            "id": message.EntryID,
            "subject": message.Subject,
            "sender": message.SenderEmailAddress,
            "received": received,
            "body": message.Body,
            "attachments": message.Attachments,   # a collection, handled later
        })
    return results

if __name__ == "__main__":
   for email in get_recent_emails(60):
    folder = save_email(email, r"C:\EmailAssistant\Output")
    print("Saved to:", folder)