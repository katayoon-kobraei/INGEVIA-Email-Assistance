from __future__ import annotations

import csv
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Iterable

from src.dehu import is_dehu_sender

HIDDEN_IDS_FILENAME = "_ui_hidden_outlook_ids.json"
VISIBILITY_CHECKED_FILENAME = "_ui_visibility_checked_ids.json"
MAIL_ITEM_CLASS = 43

# Keep legacy Other/Otros reconciliation deliberately small. New messages are
# filtered during Inbox ingestion, so reopening large numbers of already-archived
# EntryIDs every 30 minutes only adds MAPI pressure without improving normal
# routing. The newest rows are still rechecked so Outlook reclassification can
# make them visible again, while old unchecked rows migrate gradually.
RECENT_RECHECK_LIMIT = 20
LEGACY_MIGRATION_LIMIT = 10

# Outlook's Focused Inbox exposes message classification as Focused/Other.
# Classic Outlook/Exchange providers do not all expose it through exactly the
# same COM surface, so these are deliberately best-effort fallbacks. A failed
# property read simply means "unknown" and never breaks email processing.
_SENDER_SMTP_DASL = (
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F",
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001E",
)

_INFERENCE_CLASSIFICATION_DASL = (
    "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/InferenceClassification",
    "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/inferenceClassification",
)


def _safe_get(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
        return default if value is None else value
    except Exception:
        return default


def _safe_call(obj: Any, method_name: str, *args: Any) -> Any:
    try:
        method = getattr(obj, method_name)
        return method(*args)
    except Exception:
        return None


def _norm(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def hidden_folder_names() -> set[str]:
    """Folder labels that should never be shown in the Email Assistant UI.

    `HIDDEN_OUTLOOK_FOLDERS` is optional and comma/semicolon separated.
    The default covers Spanish and English Outlook installations.
    """
    raw = os.environ.get("HIDDEN_OUTLOOK_FOLDERS") or "Otros,Other"
    return {_norm(part) for part in raw.replace(";", ",").split(",") if _norm(part)}


def hide_outlook_other_enabled() -> bool:
    return (os.environ.get("HIDE_OUTLOOK_OTHER") or "true").strip().lower() == "true"


def _folder_chain_contains_hidden(item: Any) -> bool:
    wanted = hidden_folder_names()
    current = _safe_get(item, "Parent")
    visited = 0
    while current is not None and visited < 12:
        name = _norm(_safe_get(current, "Name", ""))
        if name in wanted:
            return True
        parent = _safe_get(current, "Parent")
        if parent is current:
            break
        current = parent
        visited += 1
    return False


def _inference_classification(item: Any) -> str:
    # Some providers expose a late-bound property directly.
    direct = _norm(_safe_get(item, "InferenceClassification", ""))
    if direct:
        return direct

    accessor = _safe_get(item, "PropertyAccessor")
    if accessor is None:
        return ""
    for dasl in _INFERENCE_CLASSIFICATION_DASL:
        value = _safe_call(accessor, "GetProperty", dasl)
        value = _norm(value)
        if value:
            return value
    return ""


def _item_sender_is_dehu(item: Any) -> bool:
    if is_dehu_sender(_safe_get(item, "SenderEmailAddress", "")):
        return True
    accessor = _safe_get(item, "PropertyAccessor")
    if accessor is not None:
        for dasl in _SENDER_SMTP_DASL:
            if is_dehu_sender(_safe_call(accessor, "GetProperty", dasl)):
                return True
    return False


def is_hidden_outlook_item(item: Any) -> bool:
    """True for Outlook Other/"Otros" messages that should stay out of UI.

    Supports both meanings users commonly call "Otros":
      * a real Outlook folder named Otros/Other; and
      * the Focused Inbox "Other" classification/tab.

    This is a strict exclusion: there are no sender-based exceptions.
    """
    if item is None:
        return False


    if _folder_chain_contains_hidden(item):
        return True
    if not hide_outlook_other_enabled():
        return False
    classification = _inference_classification(item)
    return classification in {"other", "otros"}


def hidden_ids_path(output_root: str | os.PathLike[str]) -> Path:
    return Path(output_root) / HIDDEN_IDS_FILENAME


def load_hidden_ids(output_root: str | os.PathLike[str]) -> set[str]:
    path = hidden_ids_path(output_root)
    if not path.is_file():
        return set()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    if isinstance(payload, list):
        return {str(value).strip() for value in payload if str(value).strip()}
    if isinstance(payload, dict):
        values = payload.get("entry_ids", [])
        if isinstance(values, list):
            return {str(value).strip() for value in values if str(value).strip()}
    return set()


def save_hidden_ids(output_root: str | os.PathLike[str], entry_ids: Iterable[str]) -> None:
    root = Path(output_root)
    root.mkdir(parents=True, exist_ok=True)
    path = hidden_ids_path(root)
    payload = sorted({str(value).strip() for value in entry_ids if str(value).strip()})
    fd, temp_name = tempfile.mkstemp(prefix="hidden_outlook_", suffix=".json", dir=str(root))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass




def _load_checked_ids(output_root: str | os.PathLike[str]) -> set[str]:
    path = Path(output_root) / VISIBILITY_CHECKED_FILENAME
    if not path.is_file():
        return set()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    if isinstance(payload, list):
        return {str(value).strip() for value in payload if str(value).strip()}
    return set()


def _save_checked_ids(output_root: str | os.PathLike[str], entry_ids: Iterable[str]) -> None:
    root = Path(output_root)
    root.mkdir(parents=True, exist_ok=True)
    path = root / VISIBILITY_CHECKED_FILENAME
    payload = sorted({str(value).strip() for value in entry_ids if str(value).strip()})
    fd, temp_name = tempfile.mkstemp(prefix="visibility_checked_", suffix=".json", dir=str(root))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(temp_path, path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


def _index_rows(output_root: str | os.PathLike[str]) -> list[dict[str, str]]:
    path = Path(output_root) / "index.csv"
    if not path.is_file():
        return []
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            return list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return []


def _resolve_item(namespace: Any, entry_id: str, store_id: str = "") -> Any:
    if store_id:
        item = _safe_call(namespace, "GetItemFromID", entry_id, store_id)
        if item is not None:
            return item
    return _safe_call(namespace, "GetItemFromID", entry_id)


def sync_hidden_ids_from_outlook(output_root: str | os.PathLike[str], namespace: Any) -> int:
    """Incrementally synchronize Outlook Other/Otros visibility state.

    v1.21 queried every archived EntryID on every pipeline run. On a mailbox
    with a long history that could itself create hundreds/thousands of MAPI
    calls. New Other/Otros messages are already excluded during ingestion, so
    only a bounded legacy migration is required here. We always re-check a
    small set of newest indexed rows and migrate a bounded number of previously-unchecked old rows.
    """
    rows = _index_rows(output_root)
    hidden = load_hidden_ids(output_root)
    checked = _load_checked_ids(output_root)


    if not rows:
        save_hidden_ids(output_root, hidden)
        return len(hidden)

    recent_rows = rows[-RECENT_RECHECK_LIMIT:]
    legacy_rows: list[dict[str, str]] = []
    for row in rows:
        entry_id = str(row.get("EntryID") or "").strip()
        if entry_id and entry_id not in checked:
            legacy_rows.append(row)
            if len(legacy_rows) >= LEGACY_MIGRATION_LIMIT:
                break

    selected: list[dict[str, str]] = []
    seen: set[str] = set()
    for row in recent_rows + legacy_rows:
        entry_id = str(row.get("EntryID") or "").strip()
        if entry_id and entry_id not in seen:
            selected.append(row)
            seen.add(entry_id)

    for row in selected:
        entry_id = str(row.get("EntryID") or "").strip()
        store_id = str(row.get("StoreID") or "").strip()
        item = _resolve_item(namespace, entry_id, store_id)
        if item is None:
            continue
        try:
            item_class = _safe_get(item, "Class")
            try:
                if item_class is not None and int(item_class) != MAIL_ITEM_CLASS:
                    checked.add(entry_id)
                    continue
            except (TypeError, ValueError):
                checked.add(entry_id)
                continue

            if is_hidden_outlook_item(item):
                hidden.add(entry_id)
            elif row in recent_rows:
                # Newest records are re-checked each run, so allow a message
                # that Outlook reclassifies back to Focused to reappear.
                hidden.discard(entry_id)
            checked.add(entry_id)
        finally:
            del item

    save_hidden_ids(output_root, hidden)
    _save_checked_ids(output_root, checked)
    return len(hidden)

