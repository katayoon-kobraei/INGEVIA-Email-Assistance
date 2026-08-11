from __future__ import annotations

import csv
import os
from pathlib import Path

PRIORITY_LIST_FILENAME = "priorities.csv"
FIELDNAMES = ["EntryID", "Priority", "Date", "Subject"]


def _priority_list_path(output_root):
    return os.path.join(output_root, PRIORITY_LIST_FILENAME)


def append_to_priority_list(email, priority, output_root):
    """Append a priority using Outlook EntryID as the stable primary key.

    Repeated rows are allowed: the UI reads the newest value last. This keeps
    network-folder writes fast and avoids rewriting a growing CSV for each mail.
    """
    path = Path(_priority_list_path(output_root))
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        normalized_priority = int(priority)
    except (TypeError, ValueError):
        normalized_priority = 3
    normalized_priority = max(1, min(5, normalized_priority))

    file_exists = path.exists() and path.stat().st_size > 0
    with path.open("a", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "EntryID": str(email.get("id") or "").strip(),
            "Priority": normalized_priority,
            "Date": email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            "Subject": str(email.get("subject") or "").strip(),
        })
