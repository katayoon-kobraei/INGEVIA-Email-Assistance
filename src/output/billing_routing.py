from src.config import INTERNAL_DOMAIN, ADMINISTRACION_EMAIL, BOSS_EMAIL


def is_external_sender(email):
    """True if the sender's domain is NOT the firm's own domain."""
    sender = (email.get("sender") or "").strip().lower()
    if "@" not in sender:
        return False
    return sender.rsplit("@", 1)[-1] != INTERNAL_DOMAIN.strip().lower()


def _recipients_text(email):
    to_field = (email.get("to") or "").lower()
    cc_field = (email.get("cc") or "").lower()
    return f"{to_field} {cc_field}"


def boss_is_recipient(email):
    return BOSS_EMAIL.strip().lower() in _recipients_text(email)


def administracion_is_recipient(email):
    return ADMINISTRACION_EMAIL.strip().lower() in _recipients_text(email)