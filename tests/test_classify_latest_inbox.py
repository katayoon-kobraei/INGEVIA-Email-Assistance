"""
test_classify_latest_inbox.py

Classify-only dry run: pulls the 10 most recent real Inbox emails and
runs them through the exact same classification steps pipeline.py
uses (classify_relevance, then classify_project, then classify_address
if it applies) -- but does NOT save anything, move/scan attachments,
write to index.csv, or mark anything as processed. Safe to run
repeatedly against your real live inbox as many times as you want.

Self-contained: fetches the Inbox itself, so it does NOT require any
changes to src/ingestion/outlook_local.py -- that file stays exactly
as you already have it.

Where this goes: project root (same level as pyproject.toml).
Run: python test_classify_latest_inbox.py
"""

import win32com.client

from src.config import OUTPUT_ROOT
from src.output.project_folders import (
    list_existing_projects,
    list_existing_addresses,
    get_project_year,
    is_formal_project_code,
)
from src.classification.relevance_agent import classify_relevance
from src.classification.project_agent import classify_project
from src.classification.address_agent import classify_address

INBOX_FOLDER_ID = 6
COUNT = 10


def _get_latest_inbox_emails(count):
    """Local to this script only -- does not touch outlook_local.py."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.GetDefaultFolder(INBOX_FOLDER_ID)
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    results = []
    for message in messages:
        if len(results) >= count:
            break
        try:
            results.append({
                "id": message.EntryID,
                "subject": message.Subject,
                "sender": message.SenderEmailAddress,
                "recipient": None,
                "timestamp": message.ReceivedTime,
                "body": message.Body,
                "attachments": message.Attachments,
                "direction": "ENTRANTE",
            })
        except Exception:
            continue  # skip odd item types (meeting requests, read receipts, etc.)
    return results


def main():
    emails = _get_latest_inbox_emails(COUNT)
    print(f"Classifying the {len(emails)} most recent Inbox email(s) -- dry run, nothing will be saved.\n")

    for email in emails:
        email_year = email["timestamp"].year
        existing = list_existing_projects(OUTPUT_ROOT, [email_year])

        print("=" * 70)
        print(f"Subject: {email['subject']}")
        print(f"From:    {email['sender']}")
        print(f"Date:    {email['timestamp']}")

        try:
            relevance = classify_relevance(email)
        except Exception as e:
            print(f"  -> Relevance check FAILED: {e}")
            relevance = None

        print(f"  -> Relevant:          {relevance.is_relevant if relevance else 'unknown (check failed)'}")
        if relevance is not None and not relevance.is_relevant:
            print("  -> (would be SKIPPED -- not real client correspondence, full classification never runs)")
            continue

        try:
            match = classify_project(email, existing)
        except Exception as e:
            print(f"  -> Project classification FAILED: {e}")
            continue

        print(f"  -> Matched existing:  {match.matched_existing}")
        print(f"  -> Project folder:    {match.project_folder_name}")
        print(f"  -> Contact:           {match.contact_label}")
        print(f"  -> Topic:             {match.topic_label}")
        print(f"  -> Mentions address:  {match.mentions_specific_address}")

        if match.mentions_specific_address and is_formal_project_code(match.project_folder_name):
            company_year = get_project_year(match.project_folder_name) or email_year
            existing_addresses = list_existing_addresses(OUTPUT_ROOT, company_year, match.project_folder_name)
            try:
                addr_match = classify_address(email, existing_addresses)
                print(f"  -> Address matched:   {addr_match.matched_existing}")
                print(f"  -> Address folder:    {addr_match.address_folder_name}")
            except Exception as e:
                print(f"  -> Address classification FAILED: {e}")

    print("=" * 70)
    print("Done. Nothing was saved to disk -- this was classification only.")


if __name__ == "__main__":
    main()