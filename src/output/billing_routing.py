from src.config import INTERNAL_DOMAIN, ADMINISTRACION_EMAIL, BOSS_EMAIL


from src.config import INTERNAL_DOMAIN, ADMINISTRACION_EMAIL, BOSS_EMAIL, BANK_DOMAINS, IGNORE_SENDERS

def is_ignored_sender(email):
    """True if the sender's exact address is on the explicit ignore
    list. Separate from is_internal_sender (which only covers the
    firm's own @ingevia.com domain) -- this is for specific EXTERNAL
    addresses that should also be skipped entirely, no matter what
    the email is about."""
    sender = (email.get("sender") or "").strip().lower()
    ignored = [s.strip().lower() for s in IGNORE_SENDERS.split(",") if s.strip()]
    return sender in ignored

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


def is_internal_sender(email):
    """True if the sender's own domain is the firm's internal domain.
    Used to ignore ALL internal mail outright -- no exceptions, not
    even administracion@ingevia.com sending as herself. (This is
    separate from administracion_is_recipient(), which only ever
    checks her as a To/Cc recipient for the billing-forward path.)"""
    sender = (email.get("sender") or "").strip().lower()
    if "@" not in sender:
        return False
    return sender.rsplit("@", 1)[-1] == INTERNAL_DOMAIN.strip().lower()

def administracion_is_recipient(email):
    return ADMINISTRACION_EMAIL.strip().lower() in _recipients_text(email)