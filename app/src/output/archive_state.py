"""
archive_state.py

Tracks junk emails whose Outlook move-to-archive-folder failed (same
"message was modified" conflict flagging can hit -- see
outlook_archive.py) so it gets retried automatically on the NEXT run,
without re-running relevance classification. Mirrors flag_state.py.

Where this goes: src/output/archive_state.py (same folder as
flag_state.py, dedupe.py).
"""

import json
import os

PENDING_ARCHIVE_FILENAME = "_pending_archive.json"


def _pending_archive_path(output_root):
    return os.path.join(output_root, PENDING_ARCHIVE_FILENAME)


def load_pending_archive(output_root):
    path = _pending_archive_path(output_root)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_pending_archive(output_root, pending):
    path = _pending_archive_path(output_root)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(pending, f, indent=2, ensure_ascii=False)


def add_pending_archive(output_root, entry_id, folder_name):
    pending = load_pending_archive(output_root)
    pending[entry_id] = folder_name
    save_pending_archive(output_root, pending)


def remove_pending_archive(output_root, entry_id):
    pending = load_pending_archive(output_root)
    if entry_id in pending:
        del pending[entry_id]
        save_pending_archive(output_root, pending)