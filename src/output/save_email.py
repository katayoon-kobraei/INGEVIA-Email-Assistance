import os
import json
import shutil
import tempfile

from src.output.folder_namer import build_conversation_folder_name
from src.safety.attachment_scanner import check_attachment
from src.output.project_folders import resolve_project_relative_path

import win32com.client
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Preformatted

OL_SAVE_AS_MSG = 3  # OlSaveAsType.olMSG


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


def _save_as_pdf(text_content, folder_path):
    """Renders the same plain-text content written to email.txt into
    a simple PDF -- a print-ready copy alongside it."""
    try:
        pdf_path = os.path.join(folder_path, "email.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=LETTER)
        style = getSampleStyleSheet()["Code"]
        doc.build([Preformatted(text_content, style)])
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
    # Formal projects (already have a "{yy}-{seq}" code) and UNSORTED
    # stay top-level, in their own year. Anything else -- a bare name
    # with no code, meaning it's not a started project yet -- lands
    # inside that year's holding pen ("{yy}-000 MAILS"), reusing the
    # year it was first created in if it already exists there.
    year, relative_project_path = resolve_project_relative_path(
        output_root, project_folder_name, email["timestamp"].year
    )

    # If this email is about a specific site for a company that has
    # multiple sites, nest one level deeper into that address's own
    # folder before 03.-CORREO.
    if address_folder_name:
        relative_project_path = os.path.join(relative_project_path, address_folder_name)

    base_path = os.path.join(
        output_root, f"TRABAJOS {year}", relative_project_path, "03.-CORREO",
        email["direction"],
        build_conversation_folder_name(email, contact_label, topic_label),
    )
    folder_path = _make_unique_folder(base_path)

    text_content = (
        (f"From: {email['sender']}\n" if email["direction"] == "ENTRANTE" else f"To: {email['recipient']}\n")
        + f"Subject: {email['subject']}\nDate: {email['timestamp']}\n\n{email['body']}"
    )
    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(text_content)

    _save_as_msg(email["id"], folder_path)
    _save_as_pdf(text_content, folder_path)

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
    _save_as_pdf(text_content, folder_path)


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