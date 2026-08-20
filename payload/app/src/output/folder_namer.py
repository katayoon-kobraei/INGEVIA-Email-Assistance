import re

from src.dehu import is_dehu_email


_WINDOWS_INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_CODE_PREFIX_RE = re.compile(r'^\s*\d{2}-\d+(?:-\d+)?\s+')
_REPLY_PREFIX_RE = re.compile(r'^(?:(?:RE|RV|FW|FWD)\s*:\s*)+', re.IGNORECASE)


def normalize_plenergy_plainco_name(name):
    """PLAINCO is just Plenergy's technical intermediary (see
    is_plenergy_sender in plenergy_routing.py) -- not a separate
    client -- so a newly-created folder name should never show PLAINCO
    on its own, and should never show PLENERGY and PLAINCO side by side.

    Matching is case-insensitive. Safe to call on any folder name.
    """
    if not name:
        return name

    if not re.search(r"plainco", name, re.IGNORECASE):
        return name

    if re.search(r"plenergy", name, re.IGNORECASE):
        name = re.sub(r"[\s\-_]*plainco[\s\-_]*", " ", name, flags=re.IGNORECASE)
        name = re.sub(r"-{2,}", "-", name)
        return re.sub(r"\s{2,}", " ", name).strip(" -_")

    def _replace(match):
        token = match.group(0)
        if token.isupper():
            return "PLENERGY"
        if token[0].isupper():
            return "Plenergy"
        return "plenergy"

    return re.sub(r"plainco", _replace, name, flags=re.IGNORECASE)


def _safe_component(value, max_chars=60, underscore_spaces=False):
    """Return a Windows-folder-safe, compact human-readable component."""
    value = str(value or "").strip()
    value = _WINDOWS_INVALID_CHARS_RE.sub(" ", value)
    value = re.sub(r"\s+", " ", value).strip(" .-_\t")
    if underscore_spaces:
        value = re.sub(r"\s+", "_", value)
    if len(value) > max_chars:
        value = value[:max_chars].rstrip(" .-_")
    return value


def company_label_from_folder(company_folder_name):
    """Convert a real server folder such as ``26-003 PLENERGY`` into the
    company label used in a per-email folder name (``PLENERGY``).

    For an unmatched company proposal (which has no numeric prefix), the
    proposed company name is returned as-is.
    """
    value = _CODE_PREFIX_RE.sub("", str(company_folder_name or "")).strip()
    return normalize_plenergy_plainco_name(value or "DESCONOCIDO")


def _counterpart_name(email, contact_label=""):
    """Best human name for the folder name.

    Incoming mail uses Outlook's real sender display name. If Outlook did not
    expose one, the AI contact label is a better human fallback than an SMTP
    address. Outgoing mail uses the AI contact label / recipient.
    """
    if str(email.get("direction") or "").upper() == "ENTRANTE":
        candidates = [
            email.get("sender_name"),
            contact_label,
            email.get("sender_email"),
            email.get("sender"),
        ]
    else:
        candidates = [
            contact_label,
            email.get("recipient"),
            email.get("to"),
        ]

    for candidate in candidates:
        candidate = str(candidate or "").strip()
        if not candidate:
            continue
        # Do not use the full RFC mailbox in a folder name if one slipped in.
        if "<" in candidate and ">" in candidate:
            candidate = candidate.split("<", 1)[0].strip() or candidate
        # When the only available identity is an email address, use the local
        # part rather than putting '@domain' into a human-facing folder name.
        if "@" in candidate and " " not in candidate:
            candidate = candidate.split("@", 1)[0]
        cleaned = _safe_component(candidate, max_chars=40, underscore_spaces=True)
        if cleaned:
            return cleaned
    return "DESCONOCIDO"




def extract_dehu_organismo_emisor(body):
    """Extract the issuing authority from a DEHU notification body.

    DEHU notifications normally expose a labelled ``Organismo emisor`` field.
    Keep this deterministic so folder naming adds no Gemini/API call. Supports
    both ``Organismo emisor: NAME`` and a label followed by NAME on the next
    non-empty line. If the expected field is absent, return a safe explicit
    fallback instead of inventing an authority.
    """
    text = str(body or "").replace("\r\n", "\n").replace("\r", "\n")
    # Collapse common non-breaking spaces while preserving line boundaries.
    text = text.replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]

    label_re = re.compile(
        r"^\s*organismo\s+emisor(?:\s+de\s+la\s+notificaci[oó]n)?\s*(?::|[-–—])?\s*(.*)$",
        re.IGNORECASE,
    )
    stop_re = re.compile(
        r"^(?:asunto|concepto|fecha|identificador|id\b|nif\b|nif/cif|cif\b|"
        r"titular|destinatario|referencia|c[oó]digo|enlace|acceso|expediente)\s*(?::|[-–—])?",
        re.IGNORECASE,
    )

    for index, line in enumerate(lines):
        match = label_re.match(line)
        if not match:
            continue

        candidate = match.group(1).strip(" :-–—\t")
        if candidate:
            return _safe_component(candidate, max_chars=90) or "Organismo Desconocido"

        # Some DEHU templates place the value on the next non-empty line.
        for next_line in lines[index + 1:index + 7]:
            candidate = next_line.strip(" :-–—\t")
            if not candidate:
                continue
            if stop_re.match(candidate):
                break
            return _safe_component(candidate, max_chars=90) or "Organismo Desconocido"
        break

    # Defensive fallback for bodies where line wrapping was lost but the label
    # and value remain in one continuous text segment.
    compact = re.sub(r"[ \t]+", " ", text)
    match = re.search(
        r"organismo\s+emisor(?:\s+de\s+la\s+notificaci[oó]n)?\s*(?::|[-–—])\s*([^\n\r]{2,120})",
        compact,
        re.IGNORECASE,
    )
    if match:
        candidate = match.group(1).strip(" :-–—\t")
        # If another labelled field leaked into the same line, cut before it.
        candidate = re.split(
            r"\s+(?=(?:asunto|concepto|fecha|identificador|nif|cif|titular|destinatario|referencia|c[oó]digo|expediente)\s*:)",
            candidate,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]
        cleaned = _safe_component(candidate, max_chars=90)
        if cleaned:
            return cleaned

    return "Organismo Desconocido"


def build_dehu_folder_name(email):
    """DEHU-only naming rule: Date + Notificacion + correo.gob.es + Organismo."""
    date_str = email["timestamp"].strftime("%y-%m-%d")
    organismo = extract_dehu_organismo_emisor(email.get("body"))
    return _safe_component(
        f"{date_str} Notificacion correo.gob.es {organismo}",
        max_chars=160,
    )

def build_routed_email_folder_name(email, company_folder_name, contact_label=""):
    """Boss-approved per-email folder naming convention.

    Format:
        YY-MM-DD Sender_Name Subject Company

    Example:
        26-08-07 Vanessa_Vierness Solicitud documentacion PLENERGY

    The actual email subject is used (reply/forward prefixes are removed only to
    keep the folder readable). Components are sanitized and length-limited so
    they remain safe on Windows/network shares.

    DEHU government notifications are the one naming exception and use:
        YY-MM-DD Notificacion correo.gob.es Organismo Emisor
    where Organismo Emisor is extracted locally from the email body.
    """
    if is_dehu_email(email):
        return build_dehu_folder_name(email)

    date_str = email["timestamp"].strftime("%y-%m-%d")
    sender = _counterpart_name(email, contact_label)

    subject = str(email.get("subject") or "SIN ASUNTO").strip()
    subject = _REPLY_PREFIX_RE.sub("", subject).strip() or "SIN ASUNTO"
    subject = _safe_component(subject, max_chars=65)

    company = _safe_component(company_label_from_folder(company_folder_name), max_chars=45)
    name = f"{date_str} {sender} {subject} {company}"
    name = normalize_plenergy_plainco_name(name)
    # Leave some room for deep network paths / duplicate suffixes.
    return _safe_component(name, max_chars=160)


def build_conversation_folder_name(email, contact_label, topic_label):
    """Legacy folder naming used by non-project special routes such as billing
    and department mail. Project correspondence uses
    ``build_routed_email_folder_name`` instead.
    """
    date_str = email["timestamp"].strftime("%y-%m-%d")
    return normalize_plenergy_plainco_name(f"{date_str} {contact_label}_{topic_label}")


def build_holding_pen_folder_name(email, project_folder_name, contact_label, topic_label):
    """Compatibility wrapper for older callers/metadata.

    v1.18 routes unmatched companies/projects with the same boss-approved
    ``date sender subject company`` convention used everywhere else.
    """
    return build_routed_email_folder_name(email, project_folder_name, contact_label)


def build_datetime_folder_name(email):
    """Legacy helper retained for compatibility."""
    return email["timestamp"].strftime("%Y-%m-%d_%H%M%S")
