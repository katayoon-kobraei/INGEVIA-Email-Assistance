"""
pending_list.py

Backend list of emails flagged as needing a reply -- date, sender,
subject. Separate from index.csv (which tracks filed correspondence)
since this is about what's still OPEN/awaiting action, not where
something was filed. Plain CSV next to index.csv under OUTPUT_ROOT,
so any spreadsheet tool (or a future dashboard panel) can open it
directly.

Where this goes: src/output/pending_list.py (same folder as
index_writer.py).
"""

import csv
import os

PENDING_LIST_FILENAME = "pendientes.csv"


def _pending_list_path(output_root):
    return os.path.join(output_root, PENDING_LIST_FILENAME)


def append_to_pending_list(email, output_root):
    path = _pending_list_path(output_root)
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["Date", "Sender", "Subject"])
        writer.writerow([
            email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            email.get("sender") or "",
            email["subject"],
        ])
