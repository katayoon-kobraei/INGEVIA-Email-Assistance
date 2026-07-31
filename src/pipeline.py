import datetime

import win32com.client

from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
    LOOKBACK_MINUTES,
)
from src.ingestion.outlook_local import get_recent_emails
from src.output.save_email import save_email
from src.output.dedupe import load_processed_ids, mark_processed
from src.output.outlook_flag import mark_email_processed
from src.output.outlook_archive import archive_email, copy_email
from src.output.flag_state import load_pending_flags, add_pending_flag, remove_pending_flag
from src.output.pending_copy_state import load_pending_copies, add_pending_copy, remove_pending_copy
from src.output.pending_list import append_to_pending_list
from src.classification.billing_agent import classify_billing
from src.output.billing_routing import is_external_sender, boss_is_recipient, administracion_is_recipient, is_internal_sender, is_ignored_sender
from src.output.save_email import save_email, save_billing_email, save_plenergy_fallback_email
from src.output.outlook_forward import forward_email
from src.config import BOSS_EMAIL, ADMINISTRACION_EMAIL, BILLING_OUTPUT_ROOT
from src.classification.plenergy_agent import classify_plenergy_address
from src.output.plenergy_routing import is_plenergy_sender, extract_us_codes, resolve_plenergy_folder, DO_PLENERGY_FOLDER, PLENERGY_FOLDER
from src.output.priority_list import append_to_priority_list
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_next_address_code,
    get_project_year,
    is_formal_project_code,
    company_uses_address_subfolders,
    get_holding_pen_name,
)
from src.classification.post_filing_agent import classify_post_filing
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.output.index_writer import append_to_index
from src.output.report_writer import append_to_report_log, generate_email_report_xlsx, cheap_fallback_summary
from src.output.status_page import generate_status_page
from src.output.run_state import load_last_run_time, save_last_run_time
from src.output.folder_namer import normalize_plenergy_plainco_name

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def _retry_pending_flags(outlook=None):
    """Emails whose Outlook flag failed on a previous run (usually
    because that exact email was open/selected in Outlook at the
    time) get retried here, before anything new is processed.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    if not FLAG_PROCESSED_EMAILS:
        return
    pending = load_pending_flags(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose Outlook flag failed last run...")
    for entry_id, category in list(pending.items()):
        if mark_email_processed(entry_id, category, outlook):
            remove_pending_flag(OUTPUT_ROOT, entry_id)
            print(f"  Flagged on retry: {entry_id}")


def _retry_pending_copies(outlook=None):
    """Retries pending-response emails whose copy to PENDING_FOLDER_NAME
    failed last run.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    if not CHECK_PENDING_RESPONSES:
        return
    pending = load_pending_copies(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} pending-email copy(ies) that failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if copy_email(entry_id, folder_name, outlook):
            remove_pending_copy(OUTPUT_ROOT, entry_id)
            print(f"  Copied on retry: {entry_id} -> {folder_name}")

def _mark_done(email, outlook=None):
    """Marks an email as processed locally (dedupe state) and, if
    enabled, stamps it back in Outlook so the boss can see it was
    handled. If the Outlook write fails (e.g. the email was
    open/selected at that moment), it's queued to retry next run.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        return
    ok = mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME, outlook)
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
    else:
        add_pending_flag(OUTPUT_ROOT, email["id"], PROCESSED_CATEGORY_NAME)
        print(f"  (Outlook flag failed -- will retry automatically next run)")


def _handle_post_filing_checks(email, outlook=None):
    """Runs the merged pending-response + priority check exactly once
    per filed email, in a single Gemini call (replaces the old separate
    _handle_pending_check + _handle_priority_check, which each re-sent
    the full email body independently). Priority is logged for every
    filed email; the pending-response copy/log only applies to incoming
    mail, same as before.

    outlook: an already-open MAPI namespace to reuse -- see
    mark_email_processed in outlook_flag.py for why this matters."""
    try:
        result = classify_post_filing(email)
    except Exception as e:
        print(f"Post-filing check failed for {email['subject']}: {e}")
        return

    append_to_priority_list(email, result.priority, OUTPUT_ROOT)
    print(f"  Priority: {result.priority}/5")

    if not CHECK_PENDING_RESPONSES or email.get("direction") != "ENTRANTE":
        return
    if not result.needs_response:
        return

    append_to_pending_list(email, OUTPUT_ROOT)
    ok = copy_email(email["id"], PENDING_FOLDER_NAME, outlook)
    if ok:
        remove_pending_copy(OUTPUT_ROOT, email["id"])
    else:
        add_pending_copy(OUTPUT_ROOT, email["id"], PENDING_FOLDER_NAME)
        print(f"  (Copy to pending folder failed -- will retry automatically next run)")

def run():
    ensure_output_root()

    # One shared Outlook connection for the entire run, threaded through
    # every function below instead of each one opening (and never
    # closing) its own. Repeatedly Dispatch()-ing a fresh
    # "Outlook.Application" connection from many separate functions,
    # run every few minutes all day by the scheduled task, is what
    # exhausted Outlook's internal resource pool and caused "Outlook ha
    # agotado todos los recursos compartidos" -- one connection per run
    # avoids that entirely.
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    _retry_pending_flags(outlook)
    _retry_pending_copies(outlook)
    processed = load_processed_ids(OUTPUT_ROOT)

    # How far back to look: normally just LOOKBACK_MINUTES (see
    # config.py), which comfortably covers the gap between two
    # scheduled runs during the day. But the scheduled task now only
    # runs 05:30-20:00, so the first run of the day follows an
    # overnight gap of roughly 9.5 hours, not 30 minutes -- a fixed
    # window would silently miss everything sent overnight. Instead,
    # look back to the timestamp of the last successful run (whatever
    # that gap actually was), with LOOKBACK_MINUTES as the floor and a
    # 5-minute safety margin added on top so a run that starts a
    # little late never leaves a sliver of a gap. This also
    # self-corrects for weekends, holidays, or a run that gets skipped
    # for some unrelated reason -- the next run just closes whatever
    # gap actually happened, without needing to special-case "the
    # first run of the day".
    run_started_at = datetime.datetime.now(datetime.timezone.utc)
    last_run = load_last_run_time(OUTPUT_ROOT)
    minutes_back = LOOKBACK_MINUTES
    if last_run is not None:
        gap_minutes = (run_started_at - last_run).total_seconds() / 60
        minutes_back = max(LOOKBACK_MINUTES, gap_minutes + 5)

    emails = get_recent_emails(minutes_back, outlook=outlook)
    save_last_run_time(OUTPUT_ROOT, run_started_at)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed. (looked back {minutes_back:.0f} min)")

    for email in emails:
        if email["id"] in processed:
            continue
        try:
            if is_internal_sender(email):
                mark_processed(email["id"], OUTPUT_ROOT)
                print(f"Ignored (internal domain): {email['subject']}")
                continue

            if is_ignored_sender(email):
                mark_processed(email["id"], OUTPUT_ROOT)
                print(f"Ignored (blocked sender): {email['subject']}")
                continue

            if is_external_sender(email) and boss_is_recipient(email):
                try:
                    billing = classify_billing(email)
                except Exception as e:
                    print(f"Billing check failed for {email['subject']}: {e}")
                    billing = None

                if billing is not None and billing.is_billing_related:
                    folder = save_billing_email(email, BILLING_OUTPUT_ROOT)
                    mark_processed(email["id"], OUTPUT_ROOT)
                    if administracion_is_recipient(email):
                        print(f"Saved (billing, admin already on it): {email['subject']} -> {folder}")
                    else:
                        ok = forward_email(email["id"], ADMINISTRACION_EMAIL, outlook)
                        if ok:
                            print(f"Saved (billing) and forwarded to {ADMINISTRACION_EMAIL}: {email['subject']} -> {folder}")
                        else:
                            print(f"Saved (billing) but forwarding FAILED: {email['subject']} -> {folder}")
                    continue

            if is_plenergy_sender(email):
                email_year = email["timestamp"].year
                us_codes = extract_us_codes(email)
                contact_label = "PLAINCO" if "plainco.es" in (email.get("sender") or "").lower() else "PLENERGY"

                # Fast path: resolve EVERY US code mentioned to its own
                # folder, not just the first one. Most emails mention
                # exactly one station, but some cover two at once (e.g.
                # a subject naming both US552 and US574) -- those need
                # to land in both stations' folders, not just one.
                matches = []  # list of (topic_label, project_folder_name, address_folder_name)
                for code in us_codes:
                    resolved = resolve_plenergy_folder(OUTPUT_ROOT, email_year, code)
                    if resolved:
                        project_folder_name, address_folder_name = resolved
                        matches.append((code, project_folder_name, address_folder_name))

                # Fallback: no US code matched anything on file yet --
                # let the model judge by full context (address, town,
                # nickname) instead. Only tried when NOTHING resolved
                # deterministically, same as before -- if one code out
                # of two already matched, that's good enough to skip
                # the extra Gemini call. If it also finds no match,
                # this same call already returned a proposed site name
                # + contact name, used below.
                llm_match = None
                if not matches:
                    do_candidates = list_existing_addresses(OUTPUT_ROOT, email_year, DO_PLENERGY_FOLDER)
                    project_candidates = list_existing_addresses(OUTPUT_ROOT, email_year, PLENERGY_FOLDER)
                    try:
                        llm_match = classify_plenergy_address(email, do_candidates, project_candidates)
                    except Exception as e:
                        print(f"Plenergy address classification failed for {email['subject']}: {e}")
                        llm_match = None

                    if llm_match is not None and llm_match.matched_existing:
                        llm_topic_label = us_codes[0] if len(us_codes) == 1 else "ESTACION IDENTIFICADA"
                        matches.append((llm_topic_label, llm_match.matched_folder, llm_match.address_folder_name))

                # llm_match is None whenever the deterministic US-code fast
                # path resolved it (no Gemini call happened at all) -- fall
                # back to a free, non-LLM summary in that one case only.
                row_summary = llm_match.summary if llm_match is not None else cheap_fallback_summary(email)

                saved_folders = []
                if matches:
                    # Save (and log) a copy under EVERY matched station --
                    # this is safe to do more than once for the same
                    # email: each call re-reads the .msg/PDF/attachments
                    # fresh from Outlook into its own destination folder,
                    # nothing is shared or moved between them.
                    for topic_label, project_folder_name, address_folder_name in matches:
                        folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT, address_folder_name, outlook=outlook)
                        append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT, address_folder_name)
                        append_to_report_log(email, contact_label, row_summary, folder, OUTPUT_ROOT)
                        saved_folders.append(folder)

                    matched_labels = {m[0] for m in matches}
                    still_unresolved = [c for c in us_codes if c not in matched_labels]
                    if still_unresolved:
                        codes_str = ", ".join(still_unresolved)
                        station_word = "station" if len(still_unresolved) == 1 else "stations"
                        print(f"  Note: subject also mentions {codes_str} -- filed under {len(matches)} matched folder(s) only; check manually if it also belongs under that {station_word}.")
                else:
                    site_hint = llm_match.address_folder_name.strip() if llm_match and llm_match.address_folder_name else ""
                    contact_name = llm_match.contact_name.strip() if llm_match and llm_match.contact_name else ""
                    date_str = email["timestamp"].strftime("%y-%m-%d")

                    parts = [date_str, contact_label]
                    if site_hint:
                        parts.append(site_hint)
                    if contact_name:
                        parts.append(contact_name)
                    folder_label = normalize_plenergy_plainco_name(" ".join(parts))

                    folder = save_plenergy_fallback_email(email, OUTPUT_ROOT, folder_label, outlook=outlook)
                    append_to_index(email, get_holding_pen_name(email_year), contact_label, folder_label, folder, OUTPUT_ROOT, None)
                    append_to_report_log(email, contact_label, row_summary, folder, OUTPUT_ROOT)
                    saved_folders.append(folder)

                _mark_done(email, outlook)
                if len(saved_folders) > 1:
                    print(f"Saved (Plenergy, {len(saved_folders)} stations): {email['subject']} -> " + " | ".join(saved_folders))
                else:
                    print(f"Saved (Plenergy): {email['subject']} -> {saved_folders[0]}")
                _handle_post_filing_checks(email, outlook)
                continue

            email_year = email["timestamp"].year

            existing = list_existing_projects(OUTPUT_ROOT, [email_year])
            address_folder_name = None
            summary = ""
            try:
                # classify_project now decides relevance itself -- no
                # separate pre-filter call, no separate token cost for
                # it. If it's not worth filing (junk, marketing, social
                # notifications, etc.), leave it completely untouched
                # in the Inbox: no flag, no move, no save. Staff
                # handle that manually. Still marked processed so it's
                # classified exactly once, never re-checked (and
                # re-billed) on future runs.
                match = classify_project(email, existing)

                if not match.is_relevant:
                    mark_processed(email["id"], OUTPUT_ROOT)
                    print(f"Ignored (not relevant): {email['subject']}")
                    continue

                project_folder_name = match.project_folder_name
                contact_label = match.contact_label
                topic_label = match.topic_label
                summary = match.summary

                # Address subfoldering only has a physical home for
                # formal projects now -- not-yet-formal entries are
                # flat, one-per-email folders directly under the
                # holding pen, with no company folder left to nest an
                # address under. Skip the call entirely for those
                # (saves the tokens too, not just the unused result).
                should_classify_address = False
                if is_formal_project_code(project_folder_name):
                    company_year = get_project_year(project_folder_name) or email_year
                    uses_addresses = company_uses_address_subfolders(OUTPUT_ROOT, company_year, project_folder_name)
                    should_classify_address = (
                        uses_addresses if uses_addresses is not None else match.mentions_specific_address
                    )

                if should_classify_address:
                    existing_addresses = list_existing_addresses(OUTPUT_ROOT, company_year, project_folder_name)
                    try:
                        addr_match = classify_address(email, existing_addresses, project_folder_name)
                        if addr_match.matched_existing:
                            address_folder_name = addr_match.address_folder_name
                        else:
                            addr_code = get_next_address_code(OUTPUT_ROOT, company_year, project_folder_name)
                            address_folder_name = f"{addr_code} {addr_match.address_folder_name}"
                    except Exception as e:
                        print(f"Address classification failed for {email['subject']}: {e}")
                        address_folder_name = None
            except Exception as e:
                print(f"Project classification failed for {email['subject']}: {e}")
                project_folder_name, contact_label, topic_label = "UNSORTED", "DESCONOCIDO", "SIN CLASIFICAR"
                summary = cheap_fallback_summary(email)

            folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT, address_folder_name, outlook=outlook)
            append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT, address_folder_name)
            append_to_report_log(email, contact_label, summary, folder, OUTPUT_ROOT)
            _mark_done(email, outlook)
            print(f"Saved: {email['subject']} -> {folder}")

            _handle_post_filing_checks(email, outlook)
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)
    generate_email_report_xlsx(OUTPUT_ROOT)


if __name__ == "__main__":
    run()