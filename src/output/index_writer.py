import csv
import os

def append_to_index(email, project_folder_name, contact_label, topic_label, folder_path, output_root, address_folder_name=None):
    index_path = os.path.join(output_root, "index.csv")
    file_exists = os.path.exists(index_path)

    with open(index_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow([
                "Year", "Project Folder", "Address Folder", "Direction", "Contact", "Topic",
                "Date", "Subject", "Sender/Recipient", "Attachments", "Folder Path"
            ])
        writer.writerow([
            email["timestamp"].year,
            project_folder_name,
            address_folder_name or "",
            email["direction"],
            contact_label,
            topic_label,
            email["timestamp"].strftime("%Y-%m-%d %H:%M"),
            email["subject"],
            email.get("sender") or email.get("recipient"),
            email["attachments"].Count,
            folder_path,
        ])