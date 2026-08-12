from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PlenergyAddressMatchResult
from src.config import GEMINI_MODEL, PROMPTS_DIR

PROMPT_PATH = PROMPTS_DIR / "plenergy_address_prompt.md"


def classify_plenergy_address(email, do_candidates, project_candidates):
    """Fallback for when the deterministic US-code string match finds
    nothing -- lets the model judge by full context (address, town,
    nickname, project details) whether this email is about an
    existing station in either folder. If no match, also returns a
    proposed new site name and the contact's name in the same call --
    both are needed for the fallback folder name, so there is no
    second Gemini call in the no-match case."""
    # Read fresh on every call -- see the matching comment in
    # project_agent.py for why.
    plenergy_rubric = PROMPT_PATH.read_text(encoding="utf-8")
    do_list = "\n".join(do_candidates) if do_candidates else "(ninguna todavía)"
    project_list = "\n".join(project_candidates) if project_candidates else "(ninguna todavía)"
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{plenergy_rubric}\n\n"
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