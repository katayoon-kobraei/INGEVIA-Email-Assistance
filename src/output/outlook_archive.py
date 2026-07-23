"""
outlook_archive.py

Moves a junk/irrelevant email out of the Inbox into a dedicated
subfolder (created automatically under the Inbox the first time it's
needed), so the real Inbox stays clean instead of just being flagged
in place. Moving between folders is a core IMAP feature, so this
works reliably even though Categories don't (see outlook_flag.py).

Where this goes: src/output/outlook_archive.py (same folder as
outlook_flag.py, save_email.py).
"""

import time

import win32com.client

INBOX_FOLDER_ID = 6


def _get_or_create_archive_folder(inbox, folder_name):
    for f in inbox.Folders:
        if f.Name == folder_name:
            return f
    return inbox.Folders.Add(folder_name)


def _try_copy_once(entry_id, folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    item = outlook.GetItemFromID(entry_id)
    target_folder = _get_or_create_archive_folder(inbox, folder_name)
    copied_item = item.Copy()
    copied_item.Move(target_folder)

def copy_email(entry_id, folder_name):
    """Same idea as archive_email, but leaves the original where it is
    -- puts a COPY in folder_name instead of moving the original out
    of the Inbox. Used for pending-response emails, which should stay
    visible in the Inbox as well as show up in the pending folder."""
    try:
        _try_copy_once(entry_id, folder_name)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_copy_once(entry_id, folder_name)
        return True
    except Exception as e:
        print(f"Could not copy email {entry_id} in Outlook: {e}")
        return False


def _try_archive_once(entry_id, folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    item = outlook.GetItemFromID(entry_id)
    archive_folder = _get_or_create_archive_folder(inbox, folder_name)
    item.Move(archive_folder)


def archive_email(entry_id, folder_name):
    """Re-fetches the item by EntryID and moves it into <folder_name>
    under the Inbox. Retries once after a short pause if the first
    attempt hits Outlook's "message was modified" conflict -- the
    same transient issue flagging can hit when the item is open or
    selected in Outlook at that exact moment. Never raises: a failed
    archive should not break the pipeline, it just prints a warning
    and returns False."""
    try:
        _try_archive_once(entry_id, folder_name)
        return True
    except Exception:
        pass  # first attempt failed -- likely a transient conflict, retry once

    time.sleep(1)
    try:
        _try_archive_once(entry_id, folder_name)
        return True
    except Exception as e:
        print(f"Could not archive email {entry_id} in Outlook: {e}")
        return False