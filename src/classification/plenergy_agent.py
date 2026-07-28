from pathlib import Path

from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PlenergyAddressMatchResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "plenergy_address_prompt.md"
PLENERGY_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_plenergy_address(email, do_candidates, project_candidates):
    """Fallback for when the deterministic US-code string match finds
    nothing -- lets the model judge by full context (address, town,
    nickname, project details) whether this email is about an
    existing station in either folder, rather than relying purely on
    an exact US-code text match."""
    do_list = "\n".join(do_candidates) if do_candidates else "(ninguna todavía)"
    project_list = "\n".join(project_candidates) if project_candidates else "(ninguna todavía)"
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{PLENERGY_RUBRIC}\n\n"
        f"Estaciones existentes en '26-004 DO PLENERGY' (fase de construcción):\n{do_list}\n\n"
        f"Estaciones existentes en '26-003 PLENERGY' (fase de proyecto):\n{project_list}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PlenergyAddressMatchResult,
        ),
    )
    return PlenergyAddressMatchResult.model_validate_json(response.text)