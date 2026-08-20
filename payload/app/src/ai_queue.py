from __future__ import annotations

import shutil
from pathlib import Path

from src.config import (
    OUTPUT_ROOT,
    ARCHIVE_ROOT,
    BILLING_OUTPUT_ROOT,
    ADMINISTRACION_EMAIL,
    CHECK_PENDING_RESPONSES,
    PENDING_FOLDER_NAME,
    AI_QUEUE_BATCH_SIZE,
    AI_QUEUE_RETRY_INITIAL_MINUTES,
    AI_QUEUE_MAX_RETRY_ATTEMPTS,
    GEMINI_429_COOLDOWN_MINUTES,
)
from src.gemini_client import GeminiRateLimited, GeminiTemporaryUnavailable
from src.gemini_state import circuit_status, open_circuit, clear_circuit
from src.dehu import is_dehu_email
from src.output.unprocessed_queue import (
    list_eligible,
    load_email,
    load_metadata,
    mark_attempt_failure,
    finalize_project_copy,
    finalize_billing_copy,
    remove_stage,
    recover_v125_path_review_items,
)
from src.output.index_writer import append_to_index, remove_entry_rows, mark_entry_ai_review, upsert_pending_to_index
from src.output.pending_copy_state import add_pending_copy
from src.output.pending_forward_state import add_pending_forward
from src.output.pending_list import append_to_pending_list
from src.output.priority_list import append_to_priority_list
from src.output.report_writer import append_to_report_log, generate_email_report_xlsx, cheap_fallback_summary
from src.output.status_page import generate_status_page
from src.output.billing_routing import (
    is_external_sender,
    boss_is_recipient,
    administracion_is_recipient,
)
from src.classification.billing_agent import classify_billing
from src.classification.plenergy_agent import classify_plenergy_address
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address
from src.classification.post_filing_agent import classify_post_filing
from src.output.plenergy_routing import (
    is_plenergy_sender,
    extract_us_codes,
    resolve_plenergy_folder,
    DO_PLENERGY_FOLDER,
    PLENERGY_FOLDER,
)
from src.output.project_folders import (
    list_existing_companies,
    list_existing_addresses,
    get_project_year,
    get_holding_pen_name,
    exact_existing_name,
    company_uses_address_subfolders,
)



def _post_filing_result(email):
    """Run the secondary AI check before any final archive copy is created.

    If Gemini is unavailable, the whole email stays in Unprocessed and can be
    retried cleanly without duplicate destination folders.
    """
    return classify_post_filing(email)


def _apply_post_filing(email, result) -> None:
    append_to_priority_list(email, result.priority, OUTPUT_ROOT)
    print(f"  Priority: {result.priority}/5")
    if CHECK_PENDING_RESPONSES and email.get("direction") == "ENTRANTE" and result.needs_response:
        append_to_pending_list(email, OUTPUT_ROOT)
        # The AI worker never opens Outlook. The regular 30-minute Outlook
        # ingestion session will perform this optional copy using EntryID/StoreID.
        add_pending_copy(
            OUTPUT_ROOT,
            email["id"],
            PENDING_FOLDER_NAME,
            email.get("store_id"),
        )


def _finalize_rows(email, stage, rows, summary, post_result) -> list[str]:
    """Copy the captured staging package to one or more final project paths.

    rows: iterable of dicts containing the same routing values used by
    finalize_project_copy / append_to_index.
    """
    saved: list[tuple[dict, str]] = []
    try:
        for row in rows:
            folder = finalize_project_copy(
                email,
                stage,
                row["project_folder_name"],
                row["contact_label"],
                row["topic_label"],
                ARCHIVE_ROOT,
                row.get("address_folder_name"),
                row.get("company_only", False),
                row.get("existing_company"),
            )
            saved.append((row, folder))
    except Exception:
        # Avoid duplicate final folders on the next queue retry if a network copy
        # fails halfway through a multi-destination email.
        for _row, folder in saved:
            shutil.rmtree(folder, ignore_errors=True)
        raise

    try:
        remove_entry_rows(OUTPUT_ROOT, email["id"])
        for row, folder in saved:
            append_to_index(
                email,
                row["project_folder_name"],
                row["contact_label"],
                row["topic_label"],
                folder,
                OUTPUT_ROOT,
                row.get("address_folder_name"),
                processing_status="PROCESADO",
            )
            append_to_report_log(email, row["contact_label"], summary, folder, OUTPUT_ROOT)

        _apply_post_filing(email, post_result)
        remove_stage(stage)
        return [folder for _row, folder in saved]
    except Exception:
        # Restore the staging row and remove copied archive folders if a shared
        # index/report write fails after the file copies completed. This keeps a
        # retry from creating duplicate destination folders.
        for _row, folder in saved:
            shutil.rmtree(folder, ignore_errors=True)
        try:
            upsert_pending_to_index(email, stage, OUTPUT_ROOT)
        except Exception:
            pass
        raise


def _process_one(stage: Path) -> None:
    email = load_email(stage)
    is_dehu = is_dehu_email(email)
    email_year = email["timestamp"].year

    # Billing remains the first AI decision for the same external-mail subset as
    # v1.24.  DEHU keeps its zero-extra-billing-call exception.
    if (not is_dehu) and is_external_sender(email) and boss_is_recipient(email):
        billing = classify_billing(email)
        if billing.is_billing_related:
            folder = finalize_billing_copy(email, stage, BILLING_OUTPUT_ROOT)
            remove_entry_rows(OUTPUT_ROOT, email["id"])
            remove_stage(stage)
            if not administracion_is_recipient(email):
                # Preserve the existing forwarding behavior without making the AI
                # retry worker touch Outlook. It is handled on the next regular
                # Outlook ingestion session. A state-file failure should not
                # duplicate the already-filed billing email on a later retry.
                try:
                    add_pending_forward(
                        OUTPUT_ROOT,
                        email["id"],
                        ADMINISTRACION_EMAIL,
                        email.get("store_id"),
                    )
                except Exception as forward_state_error:
                    print(f"Warning: billing email filed but forward could not be queued: {forward_state_error}")
            print(f"Processed staged billing email: {email['subject']} -> {folder}")
            return

    # Free deterministic US-code match for every sender.
    us_codes = extract_us_codes(email)
    us_code_matches = []
    for code in us_codes:
        resolved = resolve_plenergy_folder(ARCHIVE_ROOT, email_year, code)
        if resolved:
            project_folder_name, address_folder_name = resolved
            us_code_matches.append((code, project_folder_name, address_folder_name))

    if us_code_matches:
        contact_label = "PLAINCO" if "plainco.es" in (email.get("sender") or "").lower() else "PLENERGY"
        post_result = _post_filing_result(email)
        rows = [
            {
                "topic_label": topic,
                "project_folder_name": project,
                "address_folder_name": address,
                "contact_label": contact_label,
                "existing_company": True,
                "company_only": False,
            }
            for topic, project, address in us_code_matches
        ]
        folders = _finalize_rows(email, stage, rows, cheap_fallback_summary(email), post_result)
        print(f"Processed staged Plenergy US-code email: {email['subject']} -> " + " | ".join(folders))
        return

    if is_plenergy_sender(email):
        contact_label = "PLAINCO" if "plainco.es" in (email.get("sender") or "").lower() else "PLENERGY"
        do_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, DO_PLENERGY_FOLDER)
        project_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, PLENERGY_FOLDER)
        llm_match = classify_plenergy_address(email, do_candidates, project_candidates)
        rows = []
        if llm_match.matched_existing:
            topic_label = us_codes[0] if len(us_codes) == 1 else "ESTACION IDENTIFICADA"
            rows.append({
                "topic_label": topic_label,
                "project_folder_name": llm_match.matched_folder,
                "address_folder_name": llm_match.address_folder_name,
                "contact_label": contact_label,
                "existing_company": True,
                "company_only": False,
            })
        else:
            company_candidates = list_existing_companies(ARCHIVE_ROOT, email_year)
            plenergy_company = exact_existing_name(PLENERGY_FOLDER, company_candidates)
            topic_for_index = (
                str(llm_match.address_folder_name or "").strip()
                or (us_codes[0] if us_codes else "PROYECTO SIN IDENTIFICAR")
            )
            if plenergy_company:
                rows.append({
                    "topic_label": topic_for_index,
                    "project_folder_name": plenergy_company,
                    "address_folder_name": None,
                    "contact_label": contact_label,
                    "existing_company": True,
                    "company_only": True,
                })
            else:
                rows.append({
                    "topic_label": topic_for_index,
                    "project_folder_name": "PLENERGY",
                    "address_folder_name": None,
                    "contact_label": contact_label,
                    "existing_company": False,
                    "company_only": False,
                })
        post_result = _post_filing_result(email)
        folders = _finalize_rows(email, stage, rows, llm_match.summary or cheap_fallback_summary(email), post_result)
        print(f"Processed staged Plenergy email: {email['subject']} -> " + " | ".join(folders))
        return

    # General company-first routing. Any technical Gemini failure is allowed to
    # propagate to the queue handler; it is never converted to DESCONOCIDO.
    existing_companies = list_existing_companies(ARCHIVE_ROOT, email_year)
    company_match = classify_project(email, existing_companies)

    if not company_match.is_relevant and not is_dehu:
        # Gemini successfully decided that the message is not relevant. This is
        # a genuine classification result, not an API error, so it can leave the
        # queue permanently without being filed.
        remove_entry_rows(OUTPUT_ROOT, email["id"])
        remove_stage(stage)
        print(f"Ignored after successful AI relevance check: {email['subject']}")
        return
    if is_dehu and not company_match.is_relevant:
        print("  DEHU government-mail override: keeping email despite relevance=false.")

    contact_label = company_match.contact_label
    topic_label = company_match.topic_label
    summary = company_match.summary
    address_folder_name = None
    company_only = False
    existing_company = False

    real_company_name = exact_existing_name(company_match.project_folder_name, existing_companies)
    if real_company_name:
        existing_company = True
        project_folder_name = real_company_name
        company_year = get_project_year(project_folder_name) or email_year
        existing_projects = list_existing_addresses(ARCHIVE_ROOT, company_year, project_folder_name)
        if existing_projects:
            project_match = classify_address(email, existing_projects, project_folder_name)
            real_project_name = (
                exact_existing_name(project_match.address_folder_name, existing_projects)
                if project_match.matched_existing
                else None
            )
            if real_project_name:
                address_folder_name = real_project_name
                company_only = False
            else:
                address_folder_name = None
                company_only = True
        else:
            if company_uses_address_subfolders(ARCHIVE_ROOT, company_year, project_folder_name) is False:
                address_folder_name = None
                company_only = False
            else:
                address_folder_name = None
                company_only = True
    else:
        # This DESCONOCIDO is now only a successful-model content fallback. A
        # 429/503/timeout never reaches this line because those exceptions stay
        # in the Unprocessed queue.
        project_folder_name = (
            str(company_match.project_folder_name or "").strip()
            or str(company_match.contact_label or "").strip()
            or "DESCONOCIDO"
        )
        existing_company = False
        company_only = False

    # Same Plenergy rescue as v1.24, but a technical error also keeps the email
    # queued rather than pretending there was no match.
    rescued_via_plenergy = False
    if (not is_plenergy_sender(email)) and ((not existing_company) or company_only):
        do_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, DO_PLENERGY_FOLDER)
        project_candidates = list_existing_addresses(ARCHIVE_ROOT, email_year, PLENERGY_FOLDER)
        rescue_match = classify_plenergy_address(email, do_candidates, project_candidates)
        if rescue_match.matched_existing:
            project_folder_name = rescue_match.matched_folder
            address_folder_name = rescue_match.address_folder_name
            topic_label = us_codes[0] if len(us_codes) == 1 else "ESTACION IDENTIFICADA"
            summary = rescue_match.summary or summary
            existing_company = True
            company_only = False
            rescued_via_plenergy = True

    post_result = _post_filing_result(email)
    row = {
        "project_folder_name": project_folder_name,
        "contact_label": contact_label,
        "topic_label": topic_label,
        "address_folder_name": address_folder_name,
        "company_only": company_only,
        "existing_company": existing_company,
    }
    folders = _finalize_rows(email, stage, [row], summary or cheap_fallback_summary(email), post_result)

    if rescued_via_plenergy:
        route_kind = "Plenergy/DO PLENERGY rescue match"
    elif existing_company and address_folder_name:
        route_kind = "company + project"
    elif existing_company and company_only:
        route_kind = f"{get_holding_pen_name(email_year)} (project not found)"
    elif existing_company:
        route_kind = "company direct CORREO"
    else:
        route_kind = get_holding_pen_name(email_year)
    print(f"Processed staged [{route_kind}]: {email['subject']} -> {folders[0]}")


def run() -> int:
    # v1.26.3 automatically repairs the two known v1.25 AI_REVIEW records whose
    # final attachment path crossed legacy Windows MAX_PATH. The durable staging
    # packages are reused; Outlook is not opened.
    recovered = recover_v125_path_review_items(OUTPUT_ROOT)
    for stage in recovered:
        try:
            upsert_pending_to_index(load_email(stage), stage, OUTPUT_ROOT)
        except Exception as exc:
            print(f"Recovered queue item could not refresh its UI row: {stage.name}: {exc}")
    if recovered:
        print(f"v1.26.3 queue recovery: {len(recovered)} path-failure review item(s) returned to Pending AI.")

    is_open, until, reason = circuit_status()
    if is_open:
        print(f"Gemini circuit is in cooldown until {until.isoformat()}; queue run skipped. Last error: {reason}")
        return 0

    stages = list_eligible(OUTPUT_ROOT, AI_QUEUE_BATCH_SIZE)
    if not stages:
        print("AI queue: no eligible unprocessed emails.")
        return 0

    print(f"AI queue: processing up to {len(stages)} captured email(s) without opening Outlook.")
    completed = 0
    for stage in stages:
        try:
            meta = load_metadata(stage)
            print(f"AI queue item: {meta.get('subject') or stage.name}")
            _process_one(stage)
            completed += 1
            clear_circuit()
        except GeminiRateLimited as exc:
            meta = mark_attempt_failure(
                stage,
                str(exc),
                AI_QUEUE_RETRY_INITIAL_MINUTES,
                AI_QUEUE_MAX_RETRY_ATTEMPTS,
                terminal_on_max=False,
            )
            until = open_circuit(GEMINI_429_COOLDOWN_MINUTES, str(exc))
            print(f"Gemini 429 rate limit: stopping the whole batch. Cooldown until {until.isoformat()}.")
            break
        except GeminiTemporaryUnavailable as exc:
            meta = mark_attempt_failure(
                stage,
                str(exc),
                AI_QUEUE_RETRY_INITIAL_MINUTES,
                AI_QUEUE_MAX_RETRY_ATTEMPTS,
                terminal_on_max=False,
            )
            until = open_circuit(max(2, GEMINI_429_COOLDOWN_MINUTES // 2), str(exc))
            print(f"Gemini temporarily unavailable: stopping this batch. Cooldown until {until.isoformat()}.")
            break
        except Exception as exc:
            # A per-email malformed response or local classification issue should
            # not block unrelated queued emails. Keep this item for a spaced retry.
            meta = mark_attempt_failure(
                stage,
                str(exc),
                AI_QUEUE_RETRY_INITIAL_MINUTES,
                AI_QUEUE_MAX_RETRY_ATTEMPTS,
            )
            if meta.get("state") == "AI_REVIEW":
                mark_entry_ai_review(OUTPUT_ROOT, meta.get("id") or "")
                print(f"AI queue item moved to manual review after {meta.get('attempt_count')} failed attempts: {exc}")
            else:
                print(f"AI queue item kept for retry: {exc}")
            continue

    try:
        generate_status_page(OUTPUT_ROOT)
        generate_email_report_xlsx(OUTPUT_ROOT)
    except Exception as exc:
        print(f"Report refresh after AI queue run skipped: {exc}")
    print(f"AI queue run complete: {completed} email(s) finalized.")
    return 0
