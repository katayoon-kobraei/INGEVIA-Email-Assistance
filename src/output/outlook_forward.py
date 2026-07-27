import time
import win32com.client


def _try_forward_once(entry_id, to_address):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = outlook.GetItemFromID(entry_id)
    forward = item.Forward()
    forward.To = to_address
    forward.Send()


def forward_email(entry_id, to_address):
    """Re-fetches the item by EntryID and forwards it. Retries once
    after a short pause on Outlook's transient "message was modified"
    conflict, same pattern as flagging/archiving."""
    try:
        _try_forward_once(entry_id, to_address)
        return True
    except Exception:
        pass
    time.sleep(1)
    try:
        _try_forward_once(entry_id, to_address)
        return True
    except Exception as e:
        print(f"Could not forward email {entry_id} to {to_address}: {e}")
        return False