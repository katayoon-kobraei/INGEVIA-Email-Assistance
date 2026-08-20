from __future__ import annotations

import datetime as dt
import gc
import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path

from src.outlook_errors import is_outlook_resource_error
from src.safety.attachment_scanner import check_attachment
from src.output.folder_namer import build_conversation_folder_name
from src.output.project_folders import get_holding_pen_name
from src.output.save_email import (
    OL_SAVE_AS_MSG,
    QUARANTINE_ROOT,
    _make_unique_folder,
    _save_as_pdf,
    resolve_save_destination,
)


STATE_PENDING = "PENDING_AI"
STATE_REVIEW = "AI_REVIEW"
METADATA_FILE = "metadata.json"
EMAIL_SNAPSHOT_FILE = "email_snapshot.json"
SAFE_WINDOWS_ARCHIVE_PATH = 240
V126_RECOVERY_TAG = "v1.26-path-recovery"


def queue_root(output_root) -> Path:
    root = Path(output_root) / "Unprocessed"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _queue_folder_name(email: dict) -> str:
    timestamp = email.get("timestamp")
    if isinstance(timestamp, dt.datetime):
        stamp = timestamp.strftime("%Y%m%d_%H%M%S")
    else:
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    entry_id = str(email.get("id") or "")
    digest = hashlib.sha1(entry_id.encode("utf-8", errors="ignore")).hexdigest()[:12]
    return f"{stamp}_{digest}"


def _serialize_email(email: dict) -> dict:
    data = {}
    for key, value in email.items():
        if isinstance(value, dt.datetime):
            data[key] = value.isoformat()
        elif isinstance(value, (str, int, float, bool)) or value is None:
            data[key] = value
        elif isinstance(value, (list, tuple)):
            data[key] = [str(v) for v in value]
        else:
            data[key] = str(value)
    return data


def _parse_timestamp(value) -> dt.datetime:
    if isinstance(value, dt.datetime):
        return value
    text = str(value or "").strip()
    if text:
        try:
            return dt.datetime.fromisoformat(text)
        except ValueError:
            pass
    return dt.datetime.now().astimezone()


def capture_email(email: dict, output_root, outlook) -> Path:
    """Capture one Outlook item exactly once into the shared staging queue.

    The live MailItem is opened one time. During that single lifetime the native
    .msg and every attachment are copied. The AI retry worker later operates only
    on these files / JSON and never calls Outlook for classification retries.
    """
    root = queue_root(output_root)
    stage = root / _queue_folder_name(email)
    if stage.exists():
        # Deterministic EntryID hash makes repeated capture idempotent.
        return stage
    stage.mkdir(parents=True, exist_ok=False)

    item = None
    attachments = None
    attachment_results: list[dict] = []
    try:
        store_id = email.get("store_id")
        item = outlook.GetItemFromID(email["id"], store_id) if store_id else outlook.GetItemFromID(email["id"])

        # Native Outlook copy.
        item.SaveAs(str(stage / "email.msg"), OL_SAVE_AS_MSG)

        # Save/scan all attachments while this same MailItem is already open.
        attachments = item.Attachments
        count = int(getattr(attachments, "Count", 0) or 0)
        with tempfile.TemporaryDirectory() as tmp_dir:
            for index in range(1, count + 1):
                attachment = attachments.Item(index)
                try:
                    filename = str(getattr(attachment, "FileName", "") or "").strip() or f"attachment_{index}"
                    tmp_path = Path(tmp_dir) / filename
                    attachment.SaveAsFile(str(tmp_path))
                    is_safe, reason = check_attachment(str(tmp_path))
                    if is_safe:
                        destination = stage / filename
                        # Protect the queue's own control/document files as well as
                        # duplicate attachment names.
                        reserved = {"email.msg", "email.txt", "email.pdf", METADATA_FILE, EMAIL_SNAPSHOT_FILE}
                        if destination.exists() or destination.name.casefold() in {name.casefold() for name in reserved}:
                            stem, suffix = destination.stem, destination.suffix
                            n = 2
                            destination = stage / f"{stem} ({n}){suffix}"
                            while destination.exists() or destination.name.casefold() in {name.casefold() for name in reserved}:
                                n += 1
                                destination = stage / f"{stem} ({n}){suffix}"
                        shutil.move(str(tmp_path), str(destination))
                        attachment_results.append({"filename": destination.name, "status": "saved", "reason": reason})
                    else:
                        Path(QUARANTINE_ROOT).mkdir(parents=True, exist_ok=True)
                        quarantine_name = f"{hashlib.sha1(str(email['id']).encode()).hexdigest()[:12]}_{filename}"
                        shutil.move(str(tmp_path), str(Path(QUARANTINE_ROOT) / quarantine_name))
                        attachment_results.append({"filename": filename, "status": "quarantined", "reason": reason})
                finally:
                    del attachment

        # Local representations need no additional Outlook calls.
        text_content = (
            (f"From: {email.get('sender') or ''}\n" if email.get("direction") == "ENTRANTE" else f"To: {email.get('recipient') or ''}\n")
            + f"Subject: {email.get('subject') or ''}\nDate: {email.get('timestamp')}\n\n{email.get('body') or ''}"
        )
        (stage / "email.txt").write_text(text_content, encoding="utf-8")
        _save_as_pdf(email, str(stage))

        snapshot = _serialize_email(email)
        snapshot["timestamp"] = _parse_timestamp(snapshot.get("timestamp")).isoformat()
        (stage / EMAIL_SNAPSHOT_FILE).write_text(
            json.dumps(snapshot, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        now = dt.datetime.now(dt.timezone.utc).isoformat()
        metadata = {
            "id": email.get("id") or "",
            "store_id": email.get("store_id") or "",
            "direction": email.get("direction") or "",
            "sender": email.get("sender"),
            "sender_name": email.get("sender_name"),
            "sender_email": email.get("sender_email"),
            "recipient": email.get("recipient"),
            "subject": email.get("subject") or "",
            "timestamp": _parse_timestamp(email.get("timestamp")).isoformat(),
            "state": STATE_PENDING,
            "attempt_count": 0,
            "transient_attempt_count": 0,
            "first_queued_at": now,
            "last_attempt_at": None,
            "next_retry_at": now,
            "last_error": "",
            "attachments": attachment_results,
        }
        (stage / METADATA_FILE).write_text(
            json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        return stage
    except Exception:
        try:
            shutil.rmtree(stage, ignore_errors=True)
        finally:
            raise
    finally:
        try:
            if attachments is not None:
                del attachments
            if item is not None:
                del item
        except Exception:
            pass
        gc.collect()


def load_email(stage_path) -> dict:
    stage = Path(stage_path)
    data = json.loads((stage / EMAIL_SNAPSHOT_FILE).read_text(encoding="utf-8"))
    data["timestamp"] = _parse_timestamp(data.get("timestamp"))
    try:
        data["attachment_count"] = int(data.get("attachment_count") or 0)
    except (TypeError, ValueError):
        data["attachment_count"] = len(data.get("attachment_names") or [])
    return data


def load_metadata(stage_path) -> dict:
    return json.loads((Path(stage_path) / METADATA_FILE).read_text(encoding="utf-8"))


def save_metadata(stage_path, metadata: dict) -> None:
    path = Path(stage_path) / METADATA_FILE
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, path)


def list_eligible(output_root, limit: int = 2) -> list[Path]:
    now = dt.datetime.now(dt.timezone.utc)
    rows: list[tuple[dt.datetime, Path]] = []
    root = queue_root(output_root)
    for stage in root.iterdir():
        if not stage.is_dir() or not (stage / METADATA_FILE).is_file():
            continue
        try:
            meta = load_metadata(stage)
            if str(meta.get("direction") or "").strip().upper() == "SALIENTE":
                # v1.26.3 is incoming-only. Preserve any historical staging copy
                # for audit/recovery, but never send it to Gemini or archive it.
                continue
            if meta.get("state") != STATE_PENDING:
                # Terminal staging folders can remain temporarily when a file was
                # open during finalization. Retry cleanup on later queue scans;
                # AI_REVIEW is intentionally retained for manual inspection.
                if meta.get("state") != STATE_REVIEW:
                    shutil.rmtree(stage, ignore_errors=True)
                continue
            next_retry = dt.datetime.fromisoformat(str(meta.get("next_retry_at") or meta.get("first_queued_at")))
            if next_retry.tzinfo is None:
                next_retry = next_retry.replace(tzinfo=dt.timezone.utc)
            if next_retry > now:
                continue
            queued = dt.datetime.fromisoformat(str(meta.get("first_queued_at") or meta.get("timestamp")))
            if queued.tzinfo is None:
                queued = queued.replace(tzinfo=dt.timezone.utc)
            rows.append((queued, stage))
        except Exception:
            continue
    rows.sort(key=lambda item: item[0])
    return [stage for _, stage in rows[: max(1, int(limit))]]


def mark_attempt_failure(
    stage_path,
    error: str,
    initial_retry_minutes: int = 8,
    max_attempts: int = 6,
    *,
    terminal_on_max: bool = True,
) -> dict:
    meta = load_metadata(stage_path)
    if terminal_on_max:
        attempts = int(meta.get("attempt_count") or 0) + 1
        meta["attempt_count"] = attempts
    else:
        # Provider-wide 429/5xx failures have their own counter so a series of
        # temporary quota errors cannot use up this email's manual-review budget.
        attempts = int(meta.get("transient_attempt_count") or 0) + 1
        meta["transient_attempt_count"] = attempts
    now = dt.datetime.now(dt.timezone.utc)
    # Exponential spacing for a repeatedly failing *same* email, while new queue
    # items remain eligible immediately. Cap at 2 hours.
    delay = min(120, max(1, int(initial_retry_minutes)) * (2 ** max(0, attempts - 1)))
    meta["last_attempt_at"] = now.isoformat()
    meta["last_error"] = str(error)[:1500]
    meta["next_retry_at"] = (now + dt.timedelta(minutes=delay)).isoformat()
    # A malformed response/local error can genuinely be specific to this one
    # email, so repeated failures eventually need manual review. A 429 or Gemini
    # service outage is NOT email-specific and must never strand the message in
    # AI_REVIEW merely because the provider was unavailable several times.
    if terminal_on_max and attempts >= max(1, int(max_attempts)):
        meta["state"] = STATE_REVIEW
    else:
        meta["state"] = STATE_PENDING
    save_metadata(stage_path, meta)
    return meta


def _compact_middle(value: str, max_chars: int) -> str:
    value = str(value or "").strip()
    if max_chars <= 0:
        return ""
    if len(value) <= max_chars:
        return value
    if max_chars <= 12:
        return value[:max_chars].rstrip()
    tail = max(8, min(24, max_chars // 4))
    head = max_chars - tail - 3
    return f"{value[:head].rstrip()}...{value[-tail:].lstrip()}"


def _fit_final_base_path(base_path: str, stage_path: Path) -> Path:
    """Keep archive paths comfortably below legacy Windows MAX_PATH.

    The two v1.25 emails that reached AI_REVIEW both failed at 260/262-character
    attachment paths. We preserve the beginning and end of the approved email
    folder name and shorten only as much as the deepest captured filename needs.
    Attachment filenames are shortened separately only if the parent path is so
    deep that folder-name compaction alone is insufficient.
    """
    base = Path(base_path)
    parent = base.parent
    label = base.name
    source_names = [
        item.name
        for item in Path(stage_path).iterdir()
        if item.name != EMAIL_SNAPSHOT_FILE
    ]
    longest = max((len(name) for name in source_names), default=20)
    # Reserve room for path separators and a duplicate suffix such as " (12)".
    max_label = SAFE_WINDOWS_ARCHIVE_PATH - len(str(parent)) - longest - 10
    max_label = max(36, min(len(label), max_label))
    if len(label) > max_label:
        label = _compact_middle(label, max_label)
    return parent / label


def _safe_archive_name(final_path: Path, source_name: str) -> str:
    """Return a deterministic archive filename that fits the final Windows path."""
    name = str(source_name or "file")
    available = SAFE_WINDOWS_ARCHIVE_PATH - len(str(final_path)) - 1
    if len(name) <= available:
        return name

    source = Path(name)
    suffix = source.suffix
    stem = source.stem or "file"
    digest = hashlib.sha1(name.encode("utf-8", errors="ignore")).hexdigest()[:8]
    reserve = len(suffix) + len(digest) + 3  # space + []
    stem_budget = max(8, available - reserve)
    compact = _compact_middle(stem, stem_budget)
    return f"{compact} [{digest}]{suffix}"


def _copy_stage_contents(stage_path: Path, final_path: Path) -> dict[str, str]:
    """Copy a captured package and return any original->archive filename mapping."""
    renamed: dict[str, str] = {}
    for source in Path(stage_path).iterdir():
        if source.name == EMAIL_SNAPSHOT_FILE:
            continue
        archive_name = _safe_archive_name(final_path, source.name)
        destination = final_path / archive_name
        if archive_name != source.name:
            renamed[source.name] = archive_name
        if source.is_dir():
            shutil.copytree(source, destination)
        else:
            shutil.copy2(source, destination)
    return renamed


def _copy_stage(stage_path: Path, final_path: Path) -> dict[str, str]:
    final_path.mkdir(parents=True, exist_ok=False)
    return _copy_stage_contents(stage_path, final_path)


def recover_v125_path_review_items(output_root) -> list[Path]:
    """Reset v1.25 path-length AI_REVIEW items for one automatic v1.26 retry.

    Only review records whose last error matches the known file/path finalization
    failure are touched. Other AI_REVIEW items remain manual-review records.
    """
    recovered: list[Path] = []
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    root = queue_root(output_root)
    markers = (
        "no such file or directory",
        "cannot find the path specified",
        "the system cannot find the path specified",
        "winerror 3",
        "filename or extension is too long",
        "winerror 206",
    )
    for stage in root.iterdir():
        if not stage.is_dir() or not (stage / METADATA_FILE).is_file():
            continue
        try:
            meta = load_metadata(stage)
        except Exception:
            continue
        if meta.get("state") != STATE_REVIEW:
            continue
        if str(meta.get("direction") or "").strip().upper() == "SALIENTE":
            continue
        if meta.get("recovery_tag") == V126_RECOVERY_TAG:
            continue
        error = str(meta.get("last_error") or "").casefold()
        if not any(marker in error for marker in markers):
            continue
        meta["recovered_from_error"] = meta.get("last_error") or ""
        meta["recovery_tag"] = V126_RECOVERY_TAG
        meta["state"] = STATE_PENDING
        meta["attempt_count"] = 0
        meta["next_retry_at"] = now
        meta["last_error"] = ""
        save_metadata(stage, meta)
        recovered.append(stage)
    return recovered


def _rewrite_final_metadata(final_path: Path, updates: dict) -> None:
    metadata_path = final_path / METADATA_FILE
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    except Exception:
        metadata = {}
    metadata.update(updates)
    metadata["state"] = "PROCESSED"
    metadata["processed_at"] = dt.datetime.now(dt.timezone.utc).isoformat()
    metadata_path.write_text(json.dumps(metadata, indent=2, ensure_ascii=False), encoding="utf-8")


def finalize_project_copy(
    email: dict,
    stage_path,
    project_folder_name,
    contact_label,
    topic_label,
    output_root,
    address_folder_name=None,
    company_only=False,
    existing_company=None,
) -> str:
    base_path = resolve_save_destination(
        email, project_folder_name, contact_label, output_root,
        address_folder_name, company_only, existing_company,
    )
    fitted_base = _fit_final_base_path(base_path, Path(stage_path))
    final_path = Path(_make_unique_folder(str(fitted_base)))
    try:
        renamed = _copy_stage_contents(Path(stage_path), final_path)
        _rewrite_final_metadata(final_path, {
            "project_folder": project_folder_name,
            "address_folder": address_folder_name,
            "company_only": bool(company_only),
            "existing_company": bool(existing_company),
            "contact_label": contact_label,
            "topic_label": topic_label,
            "archive_filename_map": renamed,
        })
        return str(final_path)
    except Exception:
        # A network/share failure halfway through a copy must not leave a partial
        # archive folder that would force the next retry to create a duplicate
        # "(2)" folder. The durable staging copy remains the source of truth.
        shutil.rmtree(final_path, ignore_errors=True)
        raise


def finalize_billing_copy(email: dict, stage_path, output_root) -> str:
    base_path = os.path.join(
        str(output_root),
        email["direction"],
        build_conversation_folder_name(email, email.get("sender") or "DESCONOCIDO", "FACTURACION"),
    )
    fitted_base = _fit_final_base_path(base_path, Path(stage_path))
    final_path = Path(_make_unique_folder(str(fitted_base)))
    try:
        renamed = _copy_stage_contents(Path(stage_path), final_path)
        _rewrite_final_metadata(final_path, {
            "category": "FACTURACION_EXTERNA",
            "archive_filename_map": renamed,
        })
        return str(final_path)
    except Exception:
        shutil.rmtree(final_path, ignore_errors=True)
        raise


def remove_stage(stage_path, terminal_state: str = "DONE") -> None:
    """Retire a staging item so it can never be processed twice.

    Mark the queue metadata terminal *before* deleting the directory. On Windows
    a user can briefly have ``email.msg`` open while the AI worker finishes; in
    that case ``rmtree`` may be unable to remove the folder immediately. Keeping
    a non-PENDING state prevents the next 2-minute scheduler run from filing the
    same email again and creating a duplicate destination folder.
    """
    stage = Path(stage_path)
    try:
        meta = load_metadata(stage)
        meta["state"] = str(terminal_state or "DONE")
        meta["completed_at"] = dt.datetime.now(dt.timezone.utc).isoformat()
        save_metadata(stage, meta)
    except Exception:
        pass
    shutil.rmtree(stage, ignore_errors=True)


def stage_count(output_root) -> tuple[int, int]:
    pending = review = 0
    root = queue_root(output_root)
    for stage in root.iterdir():
        if not stage.is_dir() or not (stage / METADATA_FILE).is_file():
            continue
        try:
            state = load_metadata(stage).get("state")
            if state == STATE_PENDING:
                pending += 1
            elif state == STATE_REVIEW:
                review += 1
        except Exception:
            pass
    return pending, review
