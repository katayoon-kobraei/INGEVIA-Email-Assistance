import csv
import os

PRIORITY_LIST_FILENAME = "priorities.csv"


def _priority_list_path(output_root):
    return os.path.join(output_root, PRIORITY_LIST_FILENAME)


def append_to_priority_list(email, priority, output_root):
    path = _priority_list_path(output_root)
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["EntryID", "Priority", "Date", "Subject"])
        writer.writerow([
            email["id"],
            priority,
            email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            email["subject"],
        ])