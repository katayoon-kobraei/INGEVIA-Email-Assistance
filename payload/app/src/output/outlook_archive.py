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

from src.outlook_errors import is_outlook_resource_error

from src.outlook_mailbox import get_target_folder

INBOX_FOLDER_ID = 6


def _get_or_create_top_level_folder(outlook, folder_name):
    """Looks for folder_name as a top-level folder at the account root
    (a sibling of Inbox, Sent Items, etc.), not nested inside Inbox --
    this is where Outlook's own built-in Archive button puts things."""
    inbox = get_target_folder(outlook, INBOX_FOLDER_ID)
    store_root = inbox.Parent
    for f in store_root.Folders:
        if f.Name == folder_name:
            return f
    return store_root.Folders.Add(folder_name)


def _try_archive_top_level_once(entry_id, folder_name, outlook=None, store_id=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = archive_folder = moved = None
    try:
        item = outlook.GetItemFromID(entry_id, store_id) if store_id else outlook.GetItemFromID(entry_id)
        archive_folder = _get_or_create_top_level_folder(outlook, folder_name)
        moved = item.Move(archive_folder)
    finally:
        if moved is not None:
            del moved
        if archive_folder is not None:
            del archive_folder
        if item is not None:
            del item


def archive_to_top_level(entry_id, folder_name, outlook=None, store_id=None):
    """Same retry behavior as archive_email, but moves into a
    top-level folder (a sibling of Inbox) instead of a subfolder
    nested inside it -- used for junk mail, so it lands in Outlook's
    real Archive folder rather than a custom subfolder.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        _try_archive_top_level_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_archive_top_level_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception as e:
        if is_outlook_resource_error(e):
            raise
        print(f"Could not archive email {entry_id} to top-level folder in Outlook: {e}")
        return False


def _get_or_create_archive_folder(inbox, folder_name):
    for f in inbox.Folders:
        if f.Name == folder_name:
            return f
    return inbox.Folders.Add(folder_name)


def _try_copy_once(entry_id, folder_name, outlook=None, store_id=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = item = target_folder = copied_item = moved_item = None
    try:
        inbox = get_target_folder(outlook, INBOX_FOLDER_ID)
        item = outlook.GetItemFromID(entry_id, store_id) if store_id else outlook.GetItemFromID(entry_id)
        target_folder = _get_or_create_archive_folder(inbox, folder_name)
        copied_item = item.Copy()
        moved_item = copied_item.Move(target_folder)
    finally:
        for obj in (moved_item, copied_item, target_folder, item, inbox):
            if obj is not None:
                try:
                    del obj
                except Exception:
                    pass

def copy_email(entry_id, folder_name, outlook=None, store_id=None):
    """Same idea as archive_email, but leaves the original where it is
    -- puts a COPY in folder_name instead of moving the original out
    of the Inbox. Used for pending-response emails, which should stay
    visible in the Inbox as well as show up in the pending folder.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        _try_copy_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_copy_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception as e:
        if is_outlook_resource_error(e):
            raise
        print(f"Could not copy email {entry_id} in Outlook: {e}")
        return False


def _try_archive_once(entry_id, folder_name, outlook=None, store_id=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = item = archive_folder = moved_item = None
    try:
        inbox = get_target_folder(outlook, INBOX_FOLDER_ID)
        item = outlook.GetItemFromID(entry_id, store_id) if store_id else outlook.GetItemFromID(entry_id)
        archive_folder = _get_or_create_archive_folder(inbox, folder_name)
        moved_item = item.Move(archive_folder)
    finally:
        for obj in (moved_item, archive_folder, item, inbox):
            if obj is not None:
                try:
                    del obj
                except Exception:
                    pass


def archive_email(entry_id, folder_name, outlook=None, store_id=None):
    """Re-fetches the item by EntryID and moves it into <folder_name>
    under the Inbox. Retries once after a short pause if the first
    attempt hits Outlook's "message was modified" conflict -- the
    same transient issue flagging can hit when the item is open or
    selected in Outlook at that exact moment. Never raises: a failed
    archive should not break the pipeline, it just prints a warning
    and returns False.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        _try_archive_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception:
        pass  # first attempt failed -- likely a transient conflict, retry once

    time.sleep(1)
    try:
        _try_archive_once(entry_id, folder_name, outlook, store_id)
        return True
    except Exception as e:
        if is_outlook_resource_error(e):
            raise
        print(f"Could not archive email {entry_id} in Outlook: {e}")
        return False