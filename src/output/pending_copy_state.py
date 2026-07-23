import json
import os

PENDING_COPY_FILENAME = "_pending_copy.json"


def _path(output_root):
    return os.path.join(output_root, PENDING_COPY_FILENAME)


def load_pending_copies(output_root):
    path = _path(output_root)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_pending_copies(output_root, pending):
    with open(_path(output_root), "w", encoding="utf-8") as f:
        json.dump(pending, f, indent=2)


def add_pending_copy(output_root, entry_id, folder_name):
    pending = load_pending_copies(output_root)
    pending[entry_id] = folder_name
    save_pending_copies(output_root, pending)


def remove_pending_copy(output_root, entry_id):
    pending = load_pending_copies(output_root)
    if entry_id in pending:
        del pending[entry_id]
        save_pending_copies(output_root, pending)