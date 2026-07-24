from pathlib import Path

from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PendingResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "pending_prompt.md"
PENDING_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_pending(email):
    """Cheap check: a short prompt with no folder structure, no
    matching rules -- just "does this email need a reply?". Meant to
    run only on already-confirmed-relevant, incoming (ENTRANTE)
    email, after filing."""
    prompt = (
        f"{PENDING_RUBRIC}\n\n"
        f"Email subject: {email['subject']}\nFrom: {email.get('sender')}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PendingResult,
        ),
    )
    return PendingResult.model_validate_json(response.text)
