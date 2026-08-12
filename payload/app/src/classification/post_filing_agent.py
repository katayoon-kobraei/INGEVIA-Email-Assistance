from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PostFilingResult
from src.config import GEMINI_MODEL, PROMPTS_DIR

PROMPT_PATH = PROMPTS_DIR / "post_filing_prompt.md"


def classify_post_filing(email):
    """Combines the old separate pending-response check and priority
    score into ONE call -- both are cheap, informational-only checks
    that run after filing and need nothing but the email itself, so
    there's no reason to pay for the email body twice."""
    # Read fresh on every call -- see the matching comment in
    # project_agent.py for why.
    post_filing_rubric = PROMPT_PATH.read_text(encoding="utf-8")
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{post_filing_rubric}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PostFilingResult,
        ),
    )
    return PostFilingResult.model_validate_json(response.text)

def fallback_post_filing_result(email):
    """Return a conservative local result when the secondary AI call fails.

    This guarantees that every filed email still receives a 1-5 priority and
    therefore a UI color. It does not replace Gemini during normal operation.
    """
    from types import SimpleNamespace
    text = f"{email.get('subject', '')} {email.get('body', '')}".casefold()
    critical_terms = (
        "urgente", "urgent", "hoy", "today", "inmediato", "immediate",
        "emergencia", "emergency", "accidente", "safety", "seguridad",
        "bloqueado", "blocked", "parado", "stopped",
    )
    high_terms = (
        "plazo", "deadline", "vence", "vencimiento", "as soon as possible",
        "cuanto antes", "necesitamos", "se requiere", "required", "reclamación",
        "complaint", "legal", "contrato", "contract",
    )
    action_terms = (
        "por favor", "please", "puedes", "podéis", "necesito", "adjuntar",
        "enviar", "confirmar", "revisar", "respuesta", "reply", "confirm",
        "send", "review", "?",
    )

    if any(term in text for term in critical_terms):
        priority = 5
    elif any(term in text for term in high_terms):
        priority = 4
    elif any(term in text for term in action_terms):
        priority = 3
    elif email.get("direction") == "SALIENTE":
        priority = 2
    else:
        priority = 2

    needs_response = (
        email.get("direction") == "ENTRANTE"
        and any(term in text for term in action_terms)
    )
    return SimpleNamespace(priority=priority, needs_response=needs_response)
