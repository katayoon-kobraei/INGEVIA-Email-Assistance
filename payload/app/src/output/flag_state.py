"""
flag_state.py

Tracks emails whose Outlook flag failed to apply -- most commonly
because that exact email was open or selected in Outlook at the
moment the script tried to write to it (Outlook blocks outside COM
changes to whatever item is currently on screen). Those emails get
retried automatically on the NEXT run, without re-classifying or
re-saving anything -- the filing already succeeded, only the visual
flag was blocked, so this just closes that gap.

Where this goes: src/output/flag_state.py (same folder as dedupe.py).
"""

import json
import os

PENDING_FLAGS_FILENAME = "_pending_flags.json"


def _pending_flags_path(output_root):
    return os.path.join(output_root, PENDING_FLAGS_FILENAME)


def load_pending_flags(output_root):
    path = _pending_flags_path(output_root)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_pending_flags(output_root, pending):
    path = _pending_flags_path(output_root)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(pending, f, indent=2, ensure_ascii=False)


def add_pending_flag(output_root, entry_id, category, store_id=None):
    """Queue a failed Outlook flag and keep the mailbox StoreID.

    Older installations stored only the category string. The retry logic stays
    backward compatible with both formats.
    """
    pending = load_pending_flags(output_root)
    pending[entry_id] = {
        "category": category or "",
        "store_id": store_id or "",
    }
    save_pending_flags(output_root, pending)


def remove_pending_flag(output_root, entry_id):
    pending = load_pending_flags(output_root)
    if entry_id in pending:
        del pending[entry_id]
        save_pending_flags(output_root, pending)