from pathlib import Path
from google.genai import types
from src.gemini_client import client
from src.classification.schemas import ProjectMatchResult

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "project_prompt.md"
PROJECT_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")

def classify_project(email, existing_projects):
    projects_list = "\n".join(existing_projects) if existing_projects else "(none yet)"
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{PROJECT_RUBRIC}\n\nExisting client/project folders:\n{projects_list}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = client.models.generate_content(
        model="gemini-3-pro",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ProjectMatchResult,
        ),
    )
    return ProjectMatchResult.model_validate_json(response.text)