import json
import os
from pathlib import Path


def _path(output_root):
    return Path(output_root) / "_pending_forwards.json"


def load_pending_forwards(output_root):
    try:
        data = json.loads(_path(output_root).read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save(output_root, data):
    path = _path(output_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, path)


def add_pending_forward(output_root, entry_id, recipient, store_id=None):
    data = load_pending_forwards(output_root)
    data[str(entry_id)] = {"recipient": recipient, "store_id": store_id or ""}
    _save(output_root, data)


def remove_pending_forward(output_root, entry_id):
    data = load_pending_forwards(output_root)
    if str(entry_id) in data:
        data.pop(str(entry_id), None)
        _save(output_root, data)
