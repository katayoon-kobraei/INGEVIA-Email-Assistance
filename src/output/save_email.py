import os
import json
import shutil
import tempfile
from xml.sax.saxutils import escape

from src.output.folder_namer import build_conversation_folder_name, build_holding_pen_folder_name
from src.safety.attachment_scanner import check_attachment
from src.output.project_folders import (
    resolve_project_relative_path,
    get_holding_pen_name,
    is_formal_project_code,
    RESERVED_TOP_LEVEL_NAMES,
)

import win32com.client
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, HRFlowable

OL_SAVE_AS_MSG = 3  # OlSaveAsType.olMSG

# Styled to match Outlook's own "Ctrl+P" print output: a bold
# counterpart heading, a ruled line, then a bold-label/value header
# block, then the body -- instead of a plain monospace text dump.
_HEADING_STYLE = ParagraphStyle("OutlookHeading", fontName="Helvetica-Bold", fontSize=13, leading=16)
_LABEL_STYLE = ParagraphStyle("OutlookLabel", fontName="Helvetica-Bold", fontSize=9.5, leading=13)
_VALUE_STYLE = ParagraphStyle("OutlookValue", fontName="Helvetica", fontSize=9.5, leading=13)
_BODY_STYLE = ParagraphStyle("OutlookBody", fontName="Helvetica", fontSize=10.5, leading=15)


def save_billing_email(email, output_root):
    """Files an external billing/procurement email directly under
    BILLING_OUTPUT_ROOT -- no project-folder resolution."""
    base_path = os.path.join(
        output_root,
        email["direction"],
        build_conversation_folder_name(email, email.get("sender") or "DESCONOCIDO", "FACTURACION"),
    )
    folder_path = _make_unique_folder(base_path)

    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(f"From: {email['sender']}\n")
        f.write(f"Subject: {email['subject']}\nDate: {email['timestamp']}\n\n{email['body']}")

    attachment_results = []
    attachments = email["attachments"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            tmp_path = os.path.join(tmp_dir, attachment.FileName)
            attachment.SaveAsFile(tmp_path)
            is_safe, reason = check_attachment(tmp_path)
            if is_safe:
                shutil.move(tmp_path, os.path.join(folder_path, attachment.FileName))
                attachment_results.append({"filename": attachment.FileName, "status": "saved", "reason": reason})
            else:
                os.makedirs(QUARANTINE_ROOT, exist_ok=True)
                shutil.move(tmp_path, os.path.join(QUARANTINE_ROOT, f"{email['id']}_{attachment.FileName}"))
                attachment_results.append({"filename": attachment.FileName, "status": "quarantined", "reason": reason})

    metadata = {
        "id": email["id"], "direction": email["direction"],
        "sender": email.get("sender"), "subject": email["subject"],
        "timestamp": email["timestamp"].isoformat(),
        "category": "FACTURACION_EXTERNA",
        "attachments": attachment_results,
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    return folder_path



def _save_as_msg(entry_id, folder_path):
    """Saves a native Outlook .msg copy alongside the .txt/.pdf
    versions -- re-fetches the live item by EntryID, same pattern
    used in outlook_flag.py/outlook_archive.py."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        item = outlook.GetItemFromID(entry_id)
        item.SaveAs(os.path.join(folder_path, "email.msg"), OL_SAVE_AS_MSG)
    except Exception as e:
        print(f"Could not save .msg copy for {entry_id}: {e}")


def _save_as_pdf(email, folder_path):
    """Renders the email into a PDF styled like Outlook's own "Ctrl+P"
    print output: a bold counterpart heading, a ruled line, a
    bold-label header block (From/To, Sent, Subject, Attachments),
    then the body -- a print-ready copy alongside email.txt/.msg."""
    try:
        pdf_path = os.path.join(folder_path, "email.pdf")
        doc = SimpleDocTemplate(
            pdf_path, pagesize=LETTER,
            topMargin=0.6 * inch, bottomMargin=0.6 * inch,
            leftMargin=0.75 * inch, rightMargin=0.75 * inch,
        )

        is_incoming = email["direction"] == "ENTRANTE"
        counterpart = (email.get("sender") if is_incoming else email.get("recipient")) or "DESCONOCIDO"

        rows = []
        if is_incoming:
            rows.append(["From:", email.get("sender") or ""])
        else:
            rows.append(["To:", email.get("recipient") or ""])
        rows.append(["Sent:", email["timestamp"].strftime("%A, %B %d, %Y %I:%M %p")])
        rows.append(["Subject:", email.get("subject") or ""])

        attachments = email.get("attachments")
        if attachments is not None and attachments.Count > 0:
            names = ", ".join(attachments.Item(i).FileName for i in range(1, attachments.Count + 1))
            rows.append(["Attachments:", names])

        table_data = [
            [Paragraph(escape(label), _LABEL_STYLE), Paragraph(escape(value), _VALUE_STYLE)]
            for label, value in rows
        ]
        table = Table(table_data, colWidths=[1.1 * inch, 5.9 * inch])
        table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))

        body_html = escape(email.get("body") or "").replace("\n", "<br/>\n")

        story = [
            Paragraph(escape(counterpart), _HEADING_STYLE),
            HRFlowable(width="100%", thickness=1, color=colors.black, spaceBefore=4, spaceAfter=10),
            table,
            Spacer(1, 16),
            Paragraph(body_html, _BODY_STYLE),
        ]

        def _draw_page_number(canvas, doc):
            canvas.saveState()
            canvas.setFont("Helvetica", 9)
            canvas.drawCentredString(doc.pagesize[0] / 2, 0.4 * inch, str(canvas.getPageNumber()))
            canvas.restoreState()

        doc.build(story, onFirstPage=_draw_page_number, onLaterPages=_draw_page_number)
    except Exception as e:
        print(f"Could not save PDF copy in {folder_path}: {e}")




QUARANTINE_ROOT = r"C:\EmailAssistant\Quarantine"


def _make_unique_folder(base_path):
    folder_path = base_path
    counter = 2
    while os.path.exists(folder_path):
        folder_path = f"{base_path} ({counter})"
        counter += 1
    os.makedirs(folder_path)
    return folder_path


def save_email(email, project_folder_name, contact_label, topic_label, output_root, address_folder_name=None):
    is_formal = is_formal_project_code(project_folder_name) or project_folder_name in RESERVED_TOP_LEVEL_NAMES

    if is_formal:
        # Formal projects (already have a "{yy}-{seq}" code) and
        # UNSORTED stay top-level, in their own year, grouped under
        # 03.-CORREO/ENTRANTE-SALIENTE as before.
        year, relative_project_path = resolve_project_relative_path(
            output_root, project_folder_name, email["timestamp"].year
        )

        # If this email is about a specific site for a company that
        # has multiple sites, nest one level deeper into that
        # address's own folder before 03.-CORREO.
        if address_folder_name:
            relative_project_path = os.path.join(relative_project_path, address_folder_name)

        base_path = os.path.join(
            output_root, f"TRABAJOS {year}", relative_project_path, "03.-CORREO",
            email["direction"],
            build_conversation_folder_name(email, contact_label, topic_label),
        )
    else:
        # Not-yet-formal project: no company subfolder, no
        # 03.-CORREO/ENTRANTE-SALIENTE nesting -- one flat,
        # fully-descriptive folder per email, directly under that
        # year's holding pen. address_folder_name is ignored here --
        # there's no company folder left for it to nest under.
        year = email["timestamp"].year
        base_path = os.path.join(
            output_root, f"TRABAJOS {year}", get_holding_pen_name(year),
            build_holding_pen_folder_name(email, project_folder_name, contact_label, topic_label),
        )

    folder_path = _make_unique_folder(base_path)

    text_content = (
        (f"From: {email['sender']}\n" if email["direction"] == "ENTRANTE" else f"To: {email['recipient']}\n")
        + f"Subject: {email['subject']}\nDate: {email['timestamp']}\n\n{email['body']}"
    )
    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(text_content)

    _save_as_msg(email["id"], folder_path)
    _save_as_pdf(email, folder_path)

    attachment_results = []
    attachments = email["attachments"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            tmp_path = os.path.join(tmp_dir, attachment.FileName)
            attachment.SaveAsFile(tmp_path)
            is_safe, reason = check_attachment(tmp_path)
            if is_safe:
                shutil.move(tmp_path, os.path.join(folder_path, attachment.FileName))
                attachment_results.append({"filename": attachment.FileName, "status": "saved", "reason": reason})
            else:
                os.makedirs(QUARANTINE_ROOT, exist_ok=True)
                shutil.move(tmp_path, os.path.join(QUARANTINE_ROOT, f"{email['id']}_{attachment.FileName}"))
                attachment_results.append({"filename": attachment.FileName, "status": "quarantined", "reason": reason})

    metadata = {
        "id": email["id"], "direction": email["direction"],
        "sender": email.get("sender"), "recipient": email.get("recipient"),
        "subject": email["subject"], "timestamp": email["timestamp"].isoformat(),
        "project_folder": project_folder_name,
        "address_folder": address_folder_name,
        "contact_label": contact_label,
        "topic_label": topic_label,
        "attachments": attachment_results,
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    return folder_path

def save_department_email(email, department_folder_name, output_root):
    """Files department mail (e.g. the secretary's domain) directly
    under OUTPUT_ROOT/DEPARTAMENTOS/<department> -- bypasses the whole
    TRABAJOS/project-folder resolution since this isn't client
    correspondence."""
    base_path = os.path.join(
        output_root, "DEPARTAMENTOS", department_folder_name,
        email["direction"],
        build_conversation_folder_name(email, email.get("sender") or "DESCONOCIDO", "DEPARTAMENTO"),
    )
    folder_path = _make_unique_folder(base_path)

    text_content = (
        (f"From: {email['sender']}\n" if email["direction"] == "ENTRANTE" else f"To: {email['recipient']}\n")
        + f"Subject: {email['subject']}\nDate: {email['timestamp']}\n\n{email['body']}"
    )
    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(text_content)

    _save_as_msg(email["id"], folder_path)
    _save_as_pdf(email, folder_path)


    attachment_results = []
    attachments = email["attachments"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            tmp_path = os.path.join(tmp_dir, attachment.FileName)
            attachment.SaveAsFile(tmp_path)
            is_safe, reason = check_attachment(tmp_path)
            if is_safe:
                shutil.move(tmp_path, os.path.join(folder_path, attachment.FileName))
                attachment_results.append({"filename": attachment.FileName, "status": "saved", "reason": reason})
            else:
                os.makedirs(QUARANTINE_ROOT, exist_ok=True)
                shutil.move(tmp_path, os.path.join(QUARANTINE_ROOT, f"{email['id']}_{attachment.FileName}"))
                attachment_results.append({"filename": attachment.FileName, "status": "quarantined", "reason": reason})

    metadata = {
        "id": email["id"], "direction": email["direction"],
        "sender": email.get("sender"), "recipient": email.get("recipient"),
        "subject": email["subject"], "timestamp": email["timestamp"].isoformat(),
        "department": department_folder_name,
        "attachments": attachment_results,
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    return folder_path


def save_plenergy_fallback_email(email, output_root, folder_label):
    """Files an unmatched Plenergy/Plainco email as a single flat
    folder directly inside that year's holding pen (e.g. "26-000
    MAILS"), named with date + company tag + site hint + contact name
    for easy manual triage -- bypasses the normal project-folder
    resolution and 03.-CORREO nesting entirely.

    This function was referenced by pipeline.py but missing from this
    file -- restored here matching the same save_email/save_billing_email
    pattern (txt + msg + pdf + attachment scan/quarantine + metadata.json)."""
    year = email["timestamp"].year
    base_path = os.path.join(output_root, f"TRABAJOS {year}", get_holding_pen_name(year), folder_label)
    folder_path = _make_unique_folder(base_path)

    text_content = (
        (f"From: {email['sender']}\n" if email["direction"] == "ENTRANTE" else f"To: {email['recipient']}\n")
        + f"Subject: {email['subject']}\nDate: {email['timestamp']}\n\n{email['body']}"
    )
    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(text_content)

    _save_as_msg(email["id"], folder_path)
    _save_as_pdf(email, folder_path)

    attachment_results = []
    attachments = email["attachments"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            tmp_path = os.path.join(tmp_dir, attachment.FileName)
            attachment.SaveAsFile(tmp_path)
            is_safe, reason = check_attachment(tmp_path)
            if is_safe:
                shutil.move(tmp_path, os.path.join(folder_path, attachment.FileName))
                attachment_results.append({"filename": attachment.FileName, "status": "saved", "reason": reason})
            else:
                os.makedirs(QUARANTINE_ROOT, exist_ok=True)
                shutil.move(tmp_path, os.path.join(QUARANTINE_ROOT, f"{email['id']}_{attachment.FileName}"))
                attachment_results.append({"filename": attachment.FileName, "status": "quarantined", "reason": reason})

    metadata = {
        "id": email["id"], "direction": email["direction"],
        "sender": email.get("sender"), "recipient": email.get("recipient"),
        "subject": email["subject"], "timestamp": email["timestamp"].isoformat(),
        "category": "PLENERGY_SIN_IDENTIFICAR",
        "attachments": attachment_results,
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)
    return folder_path