from src.config import (
    OUTPUT_ROOT,
    ensure_output_root,
    FLAG_PROCESSED_EMAILS,
    PROCESSED_CATEGORY_NAME,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
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
from src.output.plenergy_routing import is_plenergy_sender, extract_us_code, resolve_plenergy_folder, DO_PLENERGY_FOLDER, PLENERGY_FOLDER
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

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def _retry_pending_flags():
    """Emails whose Outlook flag failed on a previous run (usually
    because that exact email was open/selected in Outlook at the
    time) get retried here, before anything new is processed."""
    if not FLAG_PROCESSED_EMAILS:
        return
    pending = load_pending_flags(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} email(s) whose Outlook flag failed last run...")
    for entry_id, category in list(pending.items()):
        if mark_email_processed(entry_id, category):
            remove_pending_flag(OUTPUT_ROOT, entry_id)
            print(f"  Flagged on retry: {entry_id}")


def _retry_pending_copies():
    """Retries pending-response emails whose copy to PENDING_FOLDER_NAME
    failed last run."""
    if not CHECK_PENDING_RESPONSES:
        return
    pending = load_pending_copies(OUTPUT_ROOT)
    if not pending:
        return
    print(f"Retrying {len(pending)} pending-email copy(ies) that failed last run...")
    for entry_id, folder_name in list(pending.items()):
        if copy_email(entry_id, folder_name):
            remove_pending_copy(OUTPUT_ROOT, entry_id)
            print(f"  Copied on retry: {entry_id} -> {folder_name}")

def _mark_done(email):
    """Marks an email as processed locally (dedupe state) and, if
    enabled, stamps it back in Outlook so the boss can see it was
    handled. If the Outlook write fails (e.g. the email was
    open/selected at that moment), it's queued to retry next run."""
    mark_processed(email["id"], OUTPUT_ROOT)
    if not FLAG_PROCESSED_EMAILS:
        return
    ok = mark_email_processed(email["id"], PROCESSED_CATEGORY_NAME)
    if ok:
        remove_pending_flag(OUTPUT_ROOT, email["id"])
    else:
        add_pending_flag(OUTPUT_ROOT, email["id"], PROCESSED_CATEGORY_NAME)
        print(f"  (Outlook flag failed -- will retry automatically next run)")


def _handle_post_filing_checks(email):
    """Runs the merged pending-response + priority check exactly once
    per filed email, in a single Gemini call (replaces the old separate
    _handle_pending_check + _handle_priority_check, which each re-sent
    the full email body independently). Priority is logged for every
    filed email; the pending-response copy/log only applies to incoming
    mail, same as before."""
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
    ok = copy_email(email["id"], PENDING_FOLDER_NAME)
    if ok:
        remove_pending_copy(OUTPUT_ROOT, email["id"])
    else:
        add_pending_copy(OUTPUT_ROOT, email["id"], PENDING_FOLDER_NAME)
        print(f"  (Copy to pending folder failed -- will retry automatically next run)")

def run():
    ensure_output_root()
    _retry_pending_flags()
    _retry_pending_copies()
    processed = load_processed_ids(OUTPUT_ROOT)
    emails = get_recent_emails(20)
    print(f"Found {len(emails)} email(s), {len(processed)} already processed.")

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
                        ok = forward_email(email["id"], ADMINISTRACION_EMAIL)
                        if ok:
                            print(f"Saved (billing) and forwarded to {ADMINISTRACION_EMAIL}: {email['subject']} -> {folder}")
                        else:
                            print(f"Saved (billing) but forwarding FAILED: {email['subject']} -> {folder}")
                    continue

            if is_plenergy_sender(email):
                email_year = email["timestamp"].year
                us_code = extract_us_code(email)
                contact_label = "PLAINCO" if "plainco.es" in (email.get("sender") or "").lower() else "PLENERGY"

                # Fast path: explicit US code, cheap string match, no token cost.
                match_result = resolve_plenergy_folder(OUTPUT_ROOT, email_year, us_code) if us_code else None

                # Fallback: no US code mentioned, or the code found doesn't
                # literally match any folder name -- let the model judge by
                # full context (address, town, nickname) instead. If it
                # also finds no match, this same call already returned a
                # proposed site name + contact name, used below.
                llm_match = None
                if not match_result:
                    do_candidates = list_existing_addresses(OUTPUT_ROOT, email_year, DO_PLENERGY_FOLDER)
                    project_candidates = list_existing_addresses(OUTPUT_ROOT, email_year, PLENERGY_FOLDER)
                    try:
                        llm_match = classify_plenergy_address(email, do_candidates, project_candidates)
                    except Exception as e:
                        print(f"Plenergy address classification failed for {email['subject']}: {e}")
                        llm_match = None

                    if llm_match is not None and llm_match.matched_existing:
                        match_result = (llm_match.matched_folder, llm_match.address_folder_name)

                # llm_match is None whenever the deterministic US-code fast
                # path resolved it (no Gemini call happened at all) -- fall
                # back to a free, non-LLM summary in that one case only.
                row_summary = llm_match.summary if llm_match is not None else cheap_fallback_summary(email)

                if match_result:
                    project_folder_name, address_folder_name = match_result
                    topic_label = us_code or "ESTACION IDENTIFICADA"
                    folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT, address_folder_name)
                    append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT, address_folder_name)
                else:
                    site_hint = llm_match.address_folder_name.strip() if llm_match and llm_match.address_folder_name else ""
                    contact_name = llm_match.contact_name.strip() if llm_match and llm_match.contact_name else ""
                    date_str = email["timestamp"].strftime("%y-%m-%d")

                    parts = [date_str, contact_label]
                    if site_hint:
                        parts.append(site_hint)
                    if contact_name:
                        parts.append(contact_name)
                    folder_label = " ".join(parts)

                    folder = save_plenergy_fallback_email(email, OUTPUT_ROOT, folder_label)
                    append_to_index(email, get_holding_pen_name(email_year), contact_label, folder_label, folder, OUTPUT_ROOT, None)

                append_to_report_log(email, contact_label, row_summary, folder, OUTPUT_ROOT)
                _mark_done(email)
                print(f"Saved (Plenergy): {email['subject']} -> {folder}")
                _handle_post_filing_checks(email)
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
                        elif is_formal_project_code(project_folder_name):
                            addr_code = get_next_address_code(OUTPUT_ROOT, company_year, project_folder_name)
                            address_folder_name = f"{addr_code} {addr_match.address_folder_name}"
                        else:
                            address_folder_name = addr_match.address_folder_name
                    except Exception as e:
                        print(f"Address classification failed for {email['subject']}: {e}")
                        address_folder_name = None
            except Exception as e:
                print(f"Project classification failed for {email['subject']}: {e}")
                project_folder_name, contact_label, topic_label = "UNSORTED", "DESCONOCIDO", "SIN CLASIFICAR"
                summary = cheap_fallback_summary(email)

            folder = save_email(email, project_folder_name, contact_label, topic_label, OUTPUT_ROOT, address_folder_name)
            append_to_index(email, project_folder_name, contact_label, topic_label, folder, OUTPUT_ROOT, address_folder_name)
            append_to_report_log(email, contact_label, summary, folder, OUTPUT_ROOT)
            _mark_done(email)
            print(f"Saved: {email['subject']} -> {folder}")

            _handle_post_filing_checks(email)
        except Exception as e:
            print(f"Failed on {email['id']} ({email['subject']}): {e}")
            continue

    generate_status_page(OUTPUT_ROOT)
    generate_email_report_xlsx(OUTPUT_ROOT)


if __name__ == "__main__":
    run()