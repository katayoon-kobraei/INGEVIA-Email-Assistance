from __future__ import annotations

DEHU_SENDER = "noreply.dehu@correos.gob.es"


def normalize_address(value) -> str:
    text = str(value or "").strip().casefold()
    if "<" in text and ">" in text:
        text = text.split("<", 1)[1].split(">", 1)[0].strip()
    return text


def is_dehu_sender(value) -> bool:
    """True for the official DEHU sender used by Correos/DEHú notices."""
    return normalize_address(value) == DEHU_SENDER


def is_dehu_email(email: dict) -> bool:
    """Zero-token DEHU detection shared by ingestion, routing and naming.

    Sender identity is authoritative. The subject/body fallback preserves the
    older behavior for forwarded DEHU notices where the original sender is no
    longer the live Outlook sender.
    """
    for key in ("sender_email", "sender"):
        if is_dehu_sender(email.get(key)):
            return True
    subject = str(email.get("subject") or "")
    body = str(email.get("body") or "")
    return "dehu" in f"{subject}\n{body}".casefold()
