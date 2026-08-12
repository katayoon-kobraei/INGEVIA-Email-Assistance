from google.genai import types
from src.gemini_client import generate_content_with_retry
from src.classification.schemas import ProjectMatchResult
from src.classification.project_descriptions import enrich_candidate_list, get_company_description
from src.config import GEMINI_MODEL, PROMPTS_DIR

PROMPT_PATH = PROMPTS_DIR / "project_prompt.md"

def classify_project(email, existing_projects):
    # Attaches each candidate's company-level description (from the
    # boss's reference spreadsheet, column E) when one exists, so the
    # model has real background -- domain, typical senders, keywords
    # -- not just a bare folder name to match against. No-ops cleanly
    # if DESCRIPTIONS_XLSX_PATH isn't set: every name just passes
    # through unchanged.
    # Read fresh on every call (cheap, local disk) rather than caching
    # at import time, so an edit made in the desktop app's "Edit AI
    # prompts" screen takes effect on the very next email -- even
    # within an already-running desktop app process where Python
    # would otherwise keep serving the stale cached module-level copy.
    project_rubric = PROMPT_PATH.read_text(encoding="utf-8")
    candidates = enrich_candidate_list(existing_projects, get_company_description)
    projects_list = "\n".join(candidates) if candidates else "(none yet)"
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{project_rubric}\n\nExisting client/project folders:\n{projects_list}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ProjectMatchResult,
        ),
    )
    return ProjectMatchResult.model_validate_json(response.text)