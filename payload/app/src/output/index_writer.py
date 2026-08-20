from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path

FIELDNAMES = [
    "EntryID", "StoreID", "Year", "Project Folder", "Address Folder", "Direction",
    "Contact", "Sender Name", "Sender Email", "Topic", "Date", "Subject",
    "Sender/Recipient", "Attachments", "Folder Path", "Processing Status",
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
    _write_rows_atomic(path, rows)


def _write_rows_atomic(path: Path, rows) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
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


def _base_row(email, project_folder_name, contact_label, topic_label, folder_path, address_folder_name=None, processing_status="PROCESADO"):
    return {
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
        "Attachments": int(email.get("attachment_count") or 0),
        "Folder Path": str(folder_path),
        "Processing Status": processing_status,
    }


def append_to_index(email, project_folder_name, contact_label, topic_label, folder_path, output_root, address_folder_name=None, processing_status="PROCESADO"):
    index_path = Path(output_root) / "index.csv"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    _upgrade_schema(index_path)
    file_exists = index_path.exists() and index_path.stat().st_size > 0
    with index_path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerow(_base_row(
            email, project_folder_name, contact_label, topic_label, folder_path,
            address_folder_name, processing_status,
        ))


def remove_entry_rows(output_root, entry_id: str) -> int:
    """Remove every index row for one Outlook EntryID atomically.

    A message can have multiple final rows when it belongs to multiple Plenergy
    stations, so finalization removes the single staging row first and then
    appends every real destination row.
    """
    index_path = Path(output_root) / "index.csv"
    if not index_path.exists():
        return 0
    _upgrade_schema(index_path)
    try:
        with index_path.open("r", newline="", encoding="utf-8-sig") as handle:
            rows = list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return 0
    target = str(entry_id or "")
    kept = [row for row in rows if str(row.get("EntryID") or "") != target]
    removed = len(rows) - len(kept)
    if removed:
        _write_rows_atomic(index_path, kept)
    return removed


def upsert_pending_to_index(email, folder_path, output_root) -> None:
    """Expose a captured-but-not-yet-classified email to all UIs immediately."""
    remove_entry_rows(output_root, email.get("id") or "")
    contact = (
        email.get("sender_name")
        or email.get("sender_email")
        or email.get("sender")
        or ""
    )
    append_to_index(
        email,
        "PENDING_AI",
        contact,
        "PENDIENTE IA",
        folder_path,
        output_root,
        None,
        processing_status="PENDING_AI",
    )


def mark_entry_ai_review(output_root, entry_id: str) -> None:
    index_path = Path(output_root) / "index.csv"
    if not index_path.exists():
        return
    _upgrade_schema(index_path)
    try:
        with index_path.open("r", newline="", encoding="utf-8-sig") as handle:
            rows = list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return
    changed = False
    for row in rows:
        if str(row.get("EntryID") or "") == str(entry_id or ""):
            row["Processing Status"] = "AI_REVIEW"
            row["Project Folder"] = "AI_REVIEW"
            row["Topic"] = "REVISAR IA"
            changed = True
    if changed:
        _write_rows_atomic(index_path, rows)
