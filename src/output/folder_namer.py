import re


def normalize_plenergy_plainco_name(name):
    """PLAINCO is just Plenergy's technical intermediary (see
    is_plenergy_sender in plenergy_routing.py) -- not a separate
    client -- so a newly-created folder name should never show PLAINCO
    on its own, and should never show PLENERGY and PLAINCO side by
    side. Applied as the last step whenever a new folder name is
    composed:

      - name contains both PLAINCO and PLENERGY -> drop PLAINCO (and
        the stray separator left behind), keep PLENERGY.
      - name contains PLAINCO only -> rewrite it to PLENERGY in place,
        matching the original token's case style.
      - neither, or PLENERGY only -> name is returned unchanged.

    Matching is case-insensitive. Safe to call on any folder name,
    including ones that never mention either word -- it's then just a
    no-op. Only ever applied when composing a NEW folder name, never
    used to rename folders that already exist on disk."""
    if not name:
        return name

    if not re.search(r"plainco", name, re.IGNORECASE):
        return name

    if re.search(r"plenergy", name, re.IGNORECASE):
        # Both present -- drop PLAINCO plus one adjacent separator
        # (space, dash, underscore) on either side, so removing it
        # doesn't leave "PLENERGY-" or "PLENERGY  " behind.
        name = re.sub(r"[\s\-_]*plainco[\s\-_]*", " ", name, flags=re.IGNORECASE)
        name = re.sub(r"-{2,}", "-", name)
        return re.sub(r"\s{2,}", " ", name).strip(" -_")

    # PLAINCO only -- rename it to PLENERGY in place.
    def _replace(match):
        token = match.group(0)
        if token.isupper():
            return "PLENERGY"
        if token[0].isupper():
            return "Plenergy"
        return "plenergy"

    return re.sub(r"plainco", _replace, name, flags=re.IGNORECASE)


def build_conversation_folder_name(email, contact_label, topic_label):
    date_str = email["timestamp"].strftime("%y-%m-%d")
    return normalize_plenergy_plainco_name(f"{date_str} {contact_label}_{topic_label}")


def build_holding_pen_folder_name(email, project_folder_name, contact_label, topic_label):
    """Flat, fully self-descriptive folder name for a not-yet-formal
    project's email, filed directly under that year's holding pen
    (e.g. '26-000 MAILS') -- no company subfolder, no 03.-CORREO/
    ENTRANTE-SALIENTE nesting. Every email gets its own folder here, so
    the name has to carry everything needed to identify it from the
    listing alone: date, time, company/project name, sender, and a
    short subject hint.

    e.g. '26-07-29_11-05 ENCARNA BOSCH - E.BOSCH - PRESUPUESTO OBRA'
    """
    date_str = email["timestamp"].strftime("%y-%m-%d")
    time_str = email["timestamp"].strftime("%H-%M")
    name = f"{date_str}_{time_str} {project_folder_name} - {contact_label} - {topic_label}"
    return normalize_plenergy_plainco_name(name)


def build_datetime_folder_name(email):
    """
    Builds the innermost folder name for one specific email's exact send/receive time,
    e.g. '2026-06-10_142305'. This is where the email text, attachments, and metadata.json
    actually live.
    """
    return email["timestamp"].strftime("%Y-%m-%d_%H%M%S")