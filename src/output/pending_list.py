from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path

PENDING_LIST_FILENAME = "pendientes.csv"
FIELDNAMES = ["EntryID", "Date", "Sender", "Subject"]


def _pending_list_path(output_root):
    return os.path.join(output_root, PENDING_LIST_FILENAME)


def _upgrade_schema(path: Path) -> None:
    if not path.exists():
        return
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            if (reader.fieldnames or []) == FIELDNAMES:
                return
            rows = list(reader)
    except (OSError, csv.Error):
        return

    fd, temp_name = tempfile.mkstemp(prefix="pending_", suffix=".csv", dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with temp_path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
            writer.writeheader()
            for row in rows:
                writer.writerow({name: row.get(name, "") for name in FIELDNAMES})
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


def append_to_pending_list(email, output_root):
    path = Path(_pending_list_path(output_root))
    path.parent.mkdir(parents=True, exist_ok=True)
    _upgrade_schema(path)
    file_exists = path.exists() and path.stat().st_size > 0
    with path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "EntryID": email.get("id") or "",
            "Date": email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            "Sender": email.get("sender") or "",
            "Subject": email["subject"],
        })
