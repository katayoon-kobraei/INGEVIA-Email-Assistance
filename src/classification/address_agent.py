from pathlib import Path
from google.genai import types
from src.gemini_client import generate_content_with_retry
from src.classification.schemas import AddressMatchResult
from src.classification.project_descriptions import enrich_candidate_list, get_address_description
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "address_prompt.md"
ADDRESS_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_address(email, existing_addresses, company_folder_name=None):
    # Attaches each address candidate's site-level description (from
    # the boss's reference spreadsheet, column F) when one exists --
    # company_folder_name is needed to look those up, since
    # descriptions are keyed by the company's own project code.
    candidates = existing_addresses
    if company_folder_name:
        candidates = enrich_candidate_list(
            existing_addresses,
            lambda addr: get_address_description(company_folder_name, addr),
        )

    addresses_list = "\n".join(candidates) if candidates else "(none yet)"
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{ADDRESS_RUBRIC}\n\nExisting address folders for this company:\n{addresses_list}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=AddressMatchResult,
        ),
    )
    return AddressMatchResult.model_validate_json(response.text)