import datetime
import win32com.client

INBOX_FOLDER_ID = 6
SENT_FOLDER_ID = 5

def _fetch_from_folder(outlook, folder_id, minutes_back, direction):
    folder = outlook.GetDefaultFolder(folder_id)
    messages = folder.Items
    time_field = "ReceivedTime" if direction == "ENTRANTE" else "SentOn"
    messages.Sort(f"[{time_field}]", True)

    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=minutes_back)
    results = []

    for message in messages:
        timestamp = getattr(message, time_field)
        if timestamp < cutoff:
            break

        was_unread = bool(message.UnRead)


        results.append({
            "id": message.EntryID,
            "subject": message.Subject,
            "sender": message.SenderEmailAddress if direction == "ENTRANTE" else None,
            "recipient": message.To if direction == "SALIENTE" else None,
            "timestamp": timestamp,
            "body": message.Body,
            "attachments": message.Attachments,
            "direction": direction,
        })

        # Reading .Body over IMAP silently marks the message read on the
        # server. We're only inspecting it here, not something the boss
        # actually opened -- put the unread/bold state back exactly as
        # it was before we touched it.
        if was_unread and not message.UnRead:
            try:
                message.UnRead = True
                message.Save()
            except Exception:
                pass  # not critical -- worst case it just looks read
            
    return results

def get_recent_emails(minutes_back: int = 30) -> list[dict]:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbound = _fetch_from_folder(outlook, INBOX_FOLDER_ID, minutes_back, "ENTRANTE")
    outbound = _fetch_from_folder(outlook, SENT_FOLDER_ID, minutes_back, "SALIENTE")
    return inbound + outbound