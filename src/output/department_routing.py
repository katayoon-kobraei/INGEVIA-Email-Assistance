import re

from src.config import DEPARTMENT_SECRETARIA_EMAIL, DEPARTMENT_SECRETARIA_FOLDER_NAME, IGNORE_DOMAINS

DEPARTMENT_ROUTES = []
if DEPARTMENT_SECRETARIA_EMAIL:
    DEPARTMENT_ROUTES.append((DEPARTMENT_SECRETARIA_EMAIL.strip().lower(), DEPARTMENT_SECRETARIA_FOLDER_NAME))

_IGNORE_DOMAIN_SET = {d.strip().lower() for d in IGNORE_DOMAINS.split(",") if d.strip()}


def _sender_address(email):
    return (email.get("sender") or "").strip().lower()


def _sender_domain(email):
    match = re.search(r"@([\w.-]+)$", _sender_address(email))
    return match.group(1).lower() if match else None


def match_department(email):
    """Returns the department folder name if this email's sender
    address exactly matches a configured route, else None."""
    if email.get("direction") != "ENTRANTE":
        return None
    sender = _sender_address(email)
    if not sender:
        return None
    for route_address, folder_name in DEPARTMENT_ROUTES:
        if sender == route_address:
            return folder_name
    return None


def should_ignore(email):
    """True if this email should be completely skipped -- no
    classification, no flagging, no filing, nothing. Checked AFTER
    match_department in the pipeline, so specific addresses on an
    ignored domain (e.g. the secretary) still get their own handling
    even though the rest of that domain is ignored."""
    if email.get("direction") != "ENTRANTE":
        return False
    domain = _sender_domain(email)
    if not domain:
        return False
    return domain in _IGNORE_DOMAIN_SET