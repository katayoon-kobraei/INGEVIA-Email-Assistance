import time

import win32com.client

# Outlook's OlFlagStatus enum. pywin32 doesn't expose these as named
# constants without early binding, so the raw value is used directly.
# olFlagMarked (2) = plain red "flagged for follow-up" flag -- this is
# the one that maps to IMAP's native \Flagged flag. olFlagComplete (1)
# is a subtler gray checkmark and is an Exchange-only concept that
# doesn't reliably show/persist on IMAP accounts, so it's avoided here.
OL_FLAG_MARKED = 2


def _try_flag_once(entry_id, category, outlook=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = outlook.GetItemFromID(entry_id)

    item.FlagStatus = OL_FLAG_MARKED

    if category:
        try:
            existing = [c.strip() for c in (item.Categories or "").split(",") if c.strip()]
            if category not in existing:
                existing.append(category)
                item.Categories = ", ".join(existing)
        except Exception:
            pass  # categories unsupported on this account type (e.g. IMAP) -- not fatal

    item.Save()


def mark_email_processed(entry_id, category=None, outlook=None):
    """Re-fetches the item by its EntryID -- works no matter which
    function originally read the email -- and stamps it with a red
    follow-up flag (reliable on IMAP and Exchange alike), plus a
    best-effort category (only visible on Exchange/Microsoft 365
    accounts). Retries once after a short pause if the first attempt
    hits Outlook's "message was modified" conflict -- this happens
    occasionally when Outlook is mid-sync on that exact item (e.g.
    IMAP marking it as read) at the same moment. Never raises: a
    flagging failure should not break the pipeline, it just prints a
    warning and returns False.

    outlook: an already-open MAPI namespace to reuse (pipeline.py
    passes one shared connection through the whole run, instead of
    every flag call opening its own -- this used to be a big source of
    Outlook's "resources exhausted" errors under long-running,
    frequent automation). If not given, connects fresh."""
    try:
        _try_flag_once(entry_id, category, outlook)
        return True
    except Exception:
        pass  # first attempt failed -- likely a transient sync conflict, retry once

    time.sleep(1)
    try:
        _try_flag_once(entry_id, category, outlook)
        return True
    except Exception as e:
        print(f"Could not flag email {entry_id} in Outlook: {e}")
        return False