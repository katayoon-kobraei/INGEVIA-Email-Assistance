import datetime
import win32com.client

INBOX_FOLDER_ID = 6  # Outlook's fixed constant for the Inbox folder


INBOX_FOLDER_ID = 6

def get_recent_emails(minutes_back: int = 30) -> list[dict]:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=minutes_back)
    results = []

    for message in messages:
        received = message.ReceivedTime  # keep it timezone-aware, don't strip it
        if received < cutoff:
            break

        results.append({
            "id": message.EntryID,
            "subject": message.Subject,
            "sender": message.SenderEmailAddress,
            "received": received,
            "body": message.Body,
            "attachments": message.Attachments,
        })

    return results