import time
import win32com.client

from src.outlook_errors import is_outlook_resource_error


def _try_forward_once(entry_id, to_address, outlook=None, store_id=None):
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = None
    forward = None
    try:
        item = outlook.GetItemFromID(entry_id, store_id) if store_id else outlook.GetItemFromID(entry_id)
        forward = item.Forward()
        forward.To = to_address
        forward.Send()
    finally:
        if forward is not None:
            del forward
        if item is not None:
            del item


def forward_email(entry_id, to_address, outlook=None, store_id=None):
    """Re-fetches the item by EntryID and forwards it. Retries once
    after a short pause on Outlook's transient "message was modified"
    conflict, same pattern as flagging/archiving.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        _try_forward_once(entry_id, to_address, outlook, store_id)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_forward_once(entry_id, to_address, outlook, store_id)
        return True
    except Exception as e:
        if is_outlook_resource_error(e):
            raise
        print(f"Could not forward email {entry_id} to {to_address}: {e}")
        return False