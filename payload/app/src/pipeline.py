from __future__ import annotations

import csv
import datetime
import gc
import json
from pathlib import Path

import win32com.client

from src.config import (
    OUTPUT_ROOT,
    ARCHIVE_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
    LOOKBACK_MINUTES,
    APP_MODE,
    TARGET_MAILBOX,
    OUTLOOK_SIDE_EFFECT_BATCH_SIZE,
)
from src.ingestion.outlook_local import get_recent_emails, get_recent_dehu_emails
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import copy_email
from src.output.outlook_forward import forward_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.pending_copy_state import load_pending_copies, remove_pending_copy
from src.output.pending_forward_state import load_pending_forwards, remove_pending_forward
from src.output.billing_routing import is_internal_sender, is_ignored_sender
from src.output.index_writer import upsert_pending_to_index
from src.output.unprocessed_queue import capture_email, queue_root, load_metadata
from src.output.outlook_visibility import sync_hidden_ids_from_outlook
from src.output.run_state import load_last_run_time, save_last_run_time
from src.output.status_page import generate_status_page
from src.outlook_errors import is_outlook_resource_error
from src.dehu import is_dehu_email



def _has_attachments(email) -> bool:
    try:
        return int(email.get("attachment_count") or 0) > 0
    except (TypeError, ValueError):
        return bool(email.get("attachment_names"))

def _index_entry_ids() -> set[str]:
    path = Path(OUTPUT_ROOT) / "index.csv"
    if not path.is_file():
        return set()
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            return {
                str(row.get("EntryID") or "").strip()
                for row in csv.DictReader(handle)
                if str(row.get("EntryID") or "").strip()
            }
    except Exception:
        return set()


def _staged_entry_ids() -> set[str]:
    ids: set[str] = set()
    try:
        for stage in queue_root(OUTPUT_ROOT).iterdir():
            if not stage.is_dir():
                continue
            try:
                entry_id = str(load_metadata(stage).get("id") or "").strip()
            except Exception:
                continue
            if entry_id:
                ids.add(entry_id)
    except Exception:
        pass
    return ids


def _v126_dehu_backfill(outlook) -> int:
    """One-time targeted recovery for DEHU emails v1.25 could miss.

    v1.25 detected DEHU only from subject/body, and Other/Otros filtering could
    run before sender identity was considered. Therefore an official message from
    noreply.dehu@correos.gob.es could be silently deduped/hidden even though the
    same sender worked on another day. v1.26 performs one narrow 10-day sender
    scan, ignores messages already present in the shared index/queue, and captures
    only genuinely missing DEHU messages. This happens in the existing Outlook
    session; the AI retry worker still never opens Outlook.
    """
    marker = Path(OUTPUT_ROOT) / "_v126_dehu_backfill_done.json"
    if marker.is_file():
        return 0

    existing = _index_entry_ids() | _staged_entry_ids()
    recovered = 0
    all_ok = True
    candidates = get_recent_dehu_emails(10 * 24 * 60, outlook=outlook)
    for email in candidates:
        entry_id = str(email.get("id") or "").strip()
        if not entry_id or entry_id in existing:
            continue
        try:
            stage = capture_email(email, OUTPUT_ROOT, outlook)
            upsert_pending_to_index(email, stage, OUTPUT_ROOT)
            _mark_captured(email, outlook)
            existing.add(entry_id)
            recovered += 1
            print(f"Recovered missed DEHU email for AI queue: {email.get('subject')} -> {stage}")
        except Exception as exc:
            if is_outlook_resource_error(exc):
                raise
            all_ok = False
            print(f"DEHU historical recovery deferred for {email.get('subject')}: {exc}")

    if not all_ok:
        print("DEHU v1.26 recovery will retry on the next Outlook run.")
        return recovered

    marker.write_text(
        json.dumps(
            {
                "completed_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                "window_days": 10,
                "recovered": recovered,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    if recovered:
        print(f"DEHU v1.26 recovery completed: {recovered} missing email(s) captured.")
    else:
        print("DEHU v1.26 recovery completed: no missing DEHU emails found.")
    return recovered


def _retry_pending_flags(outlook) -> None:
    if not FLAG_PROCESSED_EMAILS:
        return
    pending = load_pending_flags(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} Outlook flag(s)...")
    for entry_id, value in list(pending.items())[:OUTLOOK_SIDE_EFFECT_BATCH_SIZE]:
        if isinstance(value, dict):
            category = value.get("category") or PROCESSED_CATEGORY_NAME
            store_id = value.get("store_id") or None
        else:
            category = value or PROCESSED_CATEGORY_NAME
            store_id = None
        if mark_email_processed(entry_id, category, outlook, store_id):
            remove_pending_flag(OUTPUT_ROOT, entry_id)


def _retry_pending_copies(outlook) -> None:
    if not CHECK_PENDING_RESPONSES:
        return
    pending = load_pending_copies(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Applying {len(pending)} queued pending-response Outlook copy/copies...")
    for entry_id, value in list(pending.items())[:OUTLOOK_SIDE_EFFECT_BATCH_SIZE]:
        if isinstance(value, dict):
            folder_name = value.get("folder_name") or PENDING_FOLDER_NAME
            store_id = value.get("store_id") or None
        else:
            folder_name = value or PENDING_FOLDER_NAME
            store_id = None
        if copy_email(entry_id, folder_name, outlook, store_id):
            remove_pending_copy(OUTPUT_ROOT, entry_id)


def _retry_pending_forwards(outlook) -> None:
    pending = load_pending_forwards(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Applying {len(pending)} queued Outlook forward(s)...")
    for entry_id, value in list(pending.items())[:OUTLOOK_SIDE_EFFECT_BATCH_SIZE]:
        if not isinstance(value, dict):
            continue
        recipient = value.get("recipient") or ""
        store_id = value.get("store_id") or None
        if recipient and forward_email(entry_id, recipient, outlook, store_id):
            remove_pending_forward(OUTPUT_ROOT, entry_id)


def _mark_captured(email, outlook) -> None:
    """Dedupe + Outlook visual mark after the local capture is safely complete."""
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        return
    try:
        ok = mark_email_processed(
            email["id"],
            PROCESSED_CATEGORY_NAME,
            outlook,
            email.get("store_id"),
        )
    except Exception:
        # The staging copy and dedupe marker are already durable. If Outlook's
        # MAPI pool is exhausted specifically while applying the visual flag,
        # preserve that side-effect request before letting the short-lived worker
        # exit. The next normal Outlook session retries it without re-capturing
        # or re-running Gemini for this email.
        add_pending_flag(
            OUTPUT_ROOT,
            email["id"],
            PROCESSED_CATEGORY_NAME,
            email.get("store_id"),
        )
        raise
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
    else:
        add_pending_flag(
            OUTPUT_ROOT,
            email["id"],
            PROCESSED_CATEGORY_NAME,
            email.get("store_id"),
        )
        print("  Outlook flag failed; queued for next normal Outlook session.")


def run():
    """Capture Outlook email once; do no Gemini work in this process.

    v1.26 keeps Outlook as an ingestion source only. Eligible
    messages are copied to OUTPUT_ROOT/Unprocessed, shown in the UI as PENDING_AI,
    flagged in Outlook, and added to the dedupe state. A separate 2-minute worker
    reads those saved files and performs every Gemini call without Outlook.
    """
    if APP_MODE != "PROCESSOR":
        raise RuntimeError(
            "This installation is configured as VIEWER. Email processing is only "
            "available on the processing computer."
        )
    ensure_output_root()
    if not Path(ARCHIVE_ROOT).is_dir():
        raise RuntimeError(
            f"The project archive is unavailable: {ARCHIVE_ROOT}. "
            "Connect to the office network / P drive and try again."
        )

    print(f"Capturing configured Outlook mailbox: {TARGET_MAILBOX}")
    print("Gemini classification is handled separately by the staged AI queue.")

    # Keep the Outlook Application and MAPI Namespace as explicit local
    # references.  The worker is intentionally short-lived, but making both
    # lifetimes explicit lets pywin32 release them deterministically before COM
    # is uninitialized instead of relying on a temporary chained Dispatch object.
    outlook_app = None
    outlook = None
    gc.collect()
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    outlook = outlook_app.GetNamespace("MAPI")
    try:
        # These are small, explicit Outlook side-effects queued by prior AI runs.
        # They reuse this one normal Outlook session and do not re-run Gemini.
        _retry_pending_flags(outlook)
        _retry_pending_copies(outlook)
        _retry_pending_forwards(outlook)
        # Drop any temporary pywin32 wrappers created by side-effect helpers
        # before beginning mailbox enumeration.
        gc.collect()

        # One-time v1.26 repair: recover official DEHU notices that v1.25 may
        # have skipped due to Other/Otros or subject/body-only detection.
        _v126_dehu_backfill(outlook)
        gc.collect()

        processed = load_processed_ids(OUTPUT_ROOT)
        run_started_at = datetime.datetime.now(datetime.timezone.utc)
        last_run = load_last_run_time(OUTPUT_ROOT)
        minutes_back = LOOKBACK_MINUTES
        if last_run is not None:
            gap_minutes = (run_started_at - last_run).total_seconds() / 60
            minutes_back = max(LOOKBACK_MINUTES, gap_minutes + 5)

        emails = get_recent_emails(minutes_back, outlook=outlook)
        # IMPORTANT: do not advance the last-successful-run checkpoint yet.
        # The checkpoint must only move after the whole capture batch finishes.
        # Otherwise a mid-batch crash can make older unprocessed emails fall
        # outside the next normal LOOKBACK_MINUTES window.
        batch_had_capture_failures = False
        print(
            f"Found {len(emails)} recent email(s), {len(processed)} already captured. "
            f"(looked back {minutes_back:.0f} min)"
        )

        captured = 0
        for email in emails:
            try:
                # Internal INGEVIA mail is intentionally excluded from Gemini,
                # archive/index/UI processing, but it must still be visibly
                # marked in Outlook so staff can see that the assistant has
                # already inspected it.  Do this BEFORE the dedupe skip so
                # messages that were recorded by v1.26.2 without an Outlook
                # flag are repaired automatically on the next overlap scan.
                if is_internal_sender(email):
                    if email["id"] not in processed:
                        mark_processed(email["id"], OUTPUT_ROOT)
                        processed.add(email["id"])
                    if FLAG_PROCESSED_EMAILS and int(email.get("flag_status") or 0) != 2:
                        _mark_captured(email, outlook)
                    print(f"Ignored (internal domain; marked in Outlook): {email['subject']}")
                    continue

                if email["id"] in processed:
                    # Repair visual Outlook marking for messages that an older
                    # version already deduped but did not flag/category (notably
                    # no-attachment external mail before v1.26.7). This does not
                    # re-run Gemini or duplicate archive/UI records.
                    if FLAG_PROCESSED_EMAILS and int(email.get("flag_status") or 0) != 2:
                        _mark_captured(email, outlook)
                    continue

                if is_ignored_sender(email):
                    mark_processed(email["id"], OUTPUT_ROOT)
                    print(f"Ignored (blocked sender): {email['subject']}")
                    continue

                is_dehu = is_dehu_email(email)
                if not _has_attachments(email) and not is_dehu:
                    # No-attachment external messages are intentionally excluded
                    # from Gemini/archive/UI processing, but staff still need a
                    # visible Outlook indication that the assistant inspected them.
                    # Use the same durable marking path as captured emails so a
                    # temporary Outlook flag failure is queued for retry.
                    _mark_captured(email, outlook)
                    processed.add(email["id"])
                    print(f"Ignored (no attachments; marked in Outlook): {email['subject']}")
                    continue

                # One explicit Outlook item open: save .msg + attachments, then
                # release it. From this point onward Gemini retries need no Outlook.
                stage = capture_email(email, OUTPUT_ROOT, outlook)
                upsert_pending_to_index(email, stage, OUTPUT_ROOT)
                _mark_captured(email, outlook)
                captured += 1
                print(f"Captured for AI queue: {email['subject']} -> {stage}")
            except Exception as exc:
                if is_outlook_resource_error(exc):
                    raise
                batch_had_capture_failures = True
                print(f"Capture failed for {email.get('subject')}: {exc}")
                continue

        try:
            hidden_count = sync_hidden_ids_from_outlook(OUTPUT_ROOT, outlook)
            if hidden_count:
                print(f"UI filter: {hidden_count} Outlook Other/Otros email(s) hidden from the app.")
        except Exception as exc:
            if is_outlook_resource_error(exc):
                raise
            print(f"UI Other/Otros sync skipped: {exc}")
        finally:
            # Visibility reconciliation uses GetItemFromID on a bounded set of
            # archived rows. Release those wrappers before the worker exits.
            gc.collect()

        # Commit the checkpoint only after every Inbox item in this batch has been
        # handled successfully. If any capture failed, keep the previous checkpoint
        # so the next scheduled/manual run automatically widens its lookback and
        # retries the missed interval. Already-captured EntryIDs are deduplicated.
        if batch_had_capture_failures:
            print(
                "Outlook capture finished with one or more item failures; "
                "last-successful-run checkpoint was NOT advanced so the next run can recover them."
            )
        else:
            save_last_run_time(OUTPUT_ROOT, run_started_at)

        try:
            generate_status_page(OUTPUT_ROOT)
        except Exception as exc:
            print(f"Status page refresh skipped: {exc}")

        print(f"Outlook capture complete: {captured} email(s) added to the AI queue.")
    finally:
        # Release the MAPI namespace before the Application wrapper.  Do not call
        # Outlook.Application.Quit(): this is the user's interactive Outlook
        # session and the assistant must never close it.
        try:
            outlook = None
        except Exception:
            pass
        try:
            outlook_app = None
        except Exception:
            pass
        gc.collect()


if __name__ == "__main__":
    run()
