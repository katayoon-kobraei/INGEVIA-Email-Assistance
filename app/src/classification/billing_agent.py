from pathlib import Path

from google.genai import types

from src.gemini_client import generate_content_with_retry
from src.classification.schemas import BillingMatchResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "billing_prompt.md"
BILLING_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_billing(email):
    """Cheap check: is this email about invoicing, quotes, budgets,
    tenders, or state contracting? Only ever called for EXTERNAL
    senders where the boss is a recipient -- see billing_routing.py."""
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{BILLING_RUBRIC}\n\n"
        f"Email subject: {email['subject']}\nFrom: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=BillingMatchResult,
        ),
    )
    return BillingMatchResult.model_validate_json(response.text)