from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path

FIELDNAMES = [
    "EntryID", "StoreID", "Year", "Project Folder", "Address Folder", "Direction",
    "Contact", "Sender Name", "Sender Email", "Topic", "Date", "Subject",
    "Sender/Recipient", "Attachments", "Folder Path",
]


def _upgrade_schema(path: Path) -> None:
    if not path.exists():
        return
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            old_fields = reader.fieldnames or []
            if old_fields == FIELDNAMES:
                return
            rows = list(reader)
    except (OSError, csv.Error):
        return

    fd, temp_name = tempfile.mkstemp(prefix="index_", suffix=".csv", dir=str(path.parent))
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


def append_to_index(email, project_folder_name, contact_label, topic_label, folder_path, output_root, address_folder_name=None):
    index_path = Path(output_root) / "index.csv"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    _upgrade_schema(index_path)
    file_exists = index_path.exists() and index_path.stat().st_size > 0

    with index_path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "EntryID": email.get("id") or "",
            "StoreID": email.get("store_id") or "",
            "Year": email["timestamp"].year,
            "Project Folder": project_folder_name,
            "Address Folder": address_folder_name or "",
            "Direction": email["direction"],
            "Contact": contact_label,
            "Sender Name": email.get("sender_name") or "",
            "Sender Email": email.get("sender_email") or email.get("sender") or "",
            "Topic": topic_label,
            "Date": email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            "Subject": email["subject"],
            "Sender/Recipient": email.get("sender") or email.get("recipient"),
            "Attachments": email["attachments"].Count,
            "Folder Path": folder_path,
        })
