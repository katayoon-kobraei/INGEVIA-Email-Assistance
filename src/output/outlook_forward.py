import time
import win32com.client


def _try_forward_once(entry_id, to_address, outlook=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = outlook.GetItemFromID(entry_id)
    forward = item.Forward()
    forward.To = to_address
    forward.Send()


def forward_email(entry_id, to_address, outlook=None):
    """Re-fetches the item by EntryID and forwards it. Retries once
    after a short pause on Outlook's transient "message was modified"
    conflict, same pattern as flagging/archiving.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        _try_forward_once(entry_id, to_address, outlook)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_forward_once(entry_id, to_address, outlook)
        return True
    except Exception as e:
        print(f"Could not forward email {entry_id} to {to_address}: {e}")
        return False