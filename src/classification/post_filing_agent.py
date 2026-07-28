from pathlib import Path

from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import PostFilingResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "post_filing_prompt.md"
POST_FILING_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_post_filing(email):
    """Combines the old separate pending-response check and priority
    score into ONE call -- both are cheap, informational-only checks
    that run after filing and need nothing but the email itself, so
    there's no reason to pay for the email body twice."""
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{POST_FILING_RUBRIC}\n\n"
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