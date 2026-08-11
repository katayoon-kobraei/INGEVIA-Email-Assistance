from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path
from typing import Any

from src.ingestion.outlook_local import resolve_sender_identity
from src.output.index_writer import FIELDNAMES, _upgrade_schema


def _open_saved_msg(outlook: Any, folder_path: str):
    if not folder_path:
        return None
    msg_path = Path(folder_path) / "email.msg"
    if not msg_path.is_file():
        return None
    try:
        return outlook.OpenSharedItem(str(msg_path))
    except Exception:
        return None


def _open_live_item(outlook: Any, entry_id: str, store_id: str):
    if not entry_id:
        return None
    try:
        return outlook.GetItemFromID(entry_id, store_id) if store_id else outlook.GetItemFromID(entry_id)
    except Exception:
        return None


def backfill_missing_sender_names(output_root, outlook: Any) -> int:
    """Enrich old index rows that predate sender display-name persistence.

    v1.13/v1.14 can already display ``Name <email>`` when both values exist,
    but records processed before that change only contain the SMTP address.
    The saved native ``email.msg`` is the best source for recovering the real
    Outlook display name without guessing from the address.  On the processor
    we also try the live Outlook item by EntryID as a fallback.

    Returns the number of rows updated.  Failures are intentionally non-fatal:
    one unusual or inaccessible legacy message must never stop normal mail
    processing.
    """
    index_path = Path(output_root) / "index.csv"
    if not index_path.is_file():
        return 0

    _upgrade_schema(index_path)
    try:
        with index_path.open("r", newline="", encoding="utf-8-sig") as handle:
            rows = list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return 0

    changed = 0
    for row in rows:
        if str(row.get("Direction") or "").upper() != "ENTRANTE":
            continue
        existing_name = str(row.get("Sender Name") or "").strip()
        existing_email = str(row.get("Sender Email") or row.get("Sender/Recipient") or "").strip()
        if existing_name and existing_email:
            continue

        item = _open_saved_msg(outlook, str(row.get("Folder Path") or ""))
        if item is None:
            item = _open_live_item(
                outlook,
                str(row.get("EntryID") or "").strip(),
                str(row.get("StoreID") or "").strip(),
            )
        if item is None:
            continue

        try:
            sender_name, sender_email = resolve_sender_identity(item)
        except Exception:
            continue

        row_changed = False
        if sender_name and not existing_name:
            row["Sender Name"] = sender_name
            row_changed = True
        if sender_email and not existing_email:
            row["Sender Email"] = sender_email
            row_changed = True
        if row_changed:
            changed += 1

    if not changed:
        return 0

    fd, temp_name = tempfile.mkstemp(prefix="index_sender_", suffix=".csv", dir=str(index_path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with temp_path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
            writer.writeheader()
            for row in rows:
                writer.writerow({name: row.get(name, "") for name in FIELDNAMES})
        os.replace(temp_path, index_path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass

    return changed
