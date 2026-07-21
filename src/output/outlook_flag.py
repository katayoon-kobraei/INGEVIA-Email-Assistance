"""
outlook_flag.py

Marks an email in Outlook itself as "seen by the AI assistant" -- a
colored category plus a completed (checkmark) flag that shows up
directly in the Inbox/Sent Items list. This is the boss's own way of
visually confirming, right inside Outlook, that every email got
looked at -- no dashboard or Estado.html needed for that check.

Applied regardless of outcome: an email that got filed into a project
folder AND one the system decided was noise (LinkedIn, marketing,
etc.) both get stamped -- the flag means "the system looked at this",
not "this was filed".

Where this goes: src/output/outlook_flag.py (same folder as
save_email.py, index_writer.py, status_page.py).
"""

import win32com.client

# Outlook's OlFlagStatus enum. pywin32 doesn't expose these as named
# constants without early binding, so the raw value is used directly.
OL_FLAG_COMPLETE = 1


def mark_email_processed(entry_id, category):
    """Re-fetches the item by its EntryID -- works no matter which
    function originally read the email -- and stamps it with a
    category plus a completed flag, then saves it back to Outlook.
    Never raises: a flagging failure should not break the pipeline,
    it just prints a warning and returns False."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = outlook.GetItemFromID(entry_id)

        existing = [c.strip() for c in (item.Categories or "").split(",") if c.strip()]
        if category not in existing:
            existing.append(category)
            item.Categories = ", ".join(existing)

        item.FlagStatus = OL_FLAG_COMPLETE
        item.Save()
        return True
    except Exception as e:
        print(f"Could not flag email {entry_id} in Outlook: {e}")
        return False