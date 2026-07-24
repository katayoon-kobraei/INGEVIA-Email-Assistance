from pathlib import Path

from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PriorityScoreResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "priority_prompt.md"
PRIORITY_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_priority(email):
    """Cheap check: a short prompt with no folder structure, no
    matching rules -- just "how urgent is this, 1 to 5?". Meant to
    run on any email that already passed the junk filter."""
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{PRIORITY_RUBRIC}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=PriorityScoreResult,
        ),
    )
    return PriorityScoreResult.model_validate_json(response.text)