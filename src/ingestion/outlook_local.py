import datetime
import win32com.client

INBOX_FOLDER_ID = 6  # Outlook's fixed constant for the Inbox folder


def get_recent_emails(minutes_back: int = 30) -> list[dict]:
    """
    Connects to the Outlook desktop app running on this PC and returns
    emails received within the last `minutes_back` minutes.

    Each email is returned as a dict with:
        id            - unique Outlook message ID (EntryID)
        subject       - subject line
        sender        - sender's email address
        received      - datetime the email arrived
        body          - plain text body
        attachments   - Outlook Attachments collection (1-indexed, use .Count / .Item(i))
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # newest first

    cutoff = datetime.datetime.now() - datetime.timedelta(minutes=minutes_back)
    results = []

    for message in messages:
        received = message.ReceivedTime.replace(tzinfo=None)
        if received < cutoff:
            break  # sorted newest-first, so nothing further back matters

        results.append({
            "id": message.EntryID,
            "subject": message.Subject,
            "sender": message.SenderEmailAddress,
            "received": received,
            "body": message.Body,
            "attachments": message.Attachments,
        })

    return results