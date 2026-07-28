"""
check_quoted_reply_ratio.py

Standalone diagnostic -- does NOT touch the pipeline or its state files.
Run this directly on the Windows PC (with Outlook open, same Python/venv
that has pywin32 installed) to see how much of each email's body is
quoted reply history versus genuinely new text.

Usage:
    python check_quoted_reply_ratio.py

Reads the most recent N Inbox items and prints, for each one:
  - total body length (chars)
  - whether a quoted-reply marker was found
  - if found: length of the "new" text before that marker, and what
    percentage of the whole body that represents

If most emails show new-text-ratio well under 100%, trimming quoted
history before sending bodies to Gemini is worth building. If it's
consistently near 100% (people mostly write fresh emails, or Outlook's
"reply" already truncates for this mailbox), skip that idea.
"""

import re
import win32com.client

INBOX_FOLDER_ID = 6
COUNT = 30  # how many recent emails to sample

# Common quoted-reply separators -- Spanish and English, since this
# mailbox gets both. Order doesn't matter; we take whichever appears
# earliest in the body.
MARKERS = [
    r"-{5,}\s*Original Message\s*-{5,}",
    r"-{5,}\s*Mensaje original\s*-{5,}",
    r"_{10,}",                                   # Outlook's underscore separator line
    r"^De:\s.*$",                                # "De: nombre@dominio.com"
    r"^From:\s.*$",
    r"^El .{0,40}escribi[oó]:",                  # "El vie, 23 jul 2026, ... escribió:"
    r"^On .{0,60}wrote:",
]

COMBINED_RE = re.compile("|".join(f"(?:{m})" for m in MARKERS), re.IGNORECASE | re.MULTILINE)


def _get_latest_inbox_emails(count):
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
                "subject": message.Subject,
                "body": message.Body or "",
            })
        except Exception:
            continue
    return results


def main():
    emails = _get_latest_inbox_emails(COUNT)
    print(f"Sampling {len(emails)} recent Inbox email(s)...\n")

    total_chars = 0
    total_new_chars = 0
    marker_hits = 0

    for email in emails:
        body = email["body"]
        total_len = len(body)
        total_chars += total_len

        match = COMBINED_RE.search(body)
        if match:
            marker_hits += 1
            new_len = match.start()
            pct = (new_len / total_len * 100) if total_len else 0
            total_new_chars += new_len
            print(f"[{email['subject'][:60]!r}] total={total_len} chars, new-text≈{new_len} chars ({pct:.0f}%)")
        else:
            total_new_chars += total_len
            print(f"[{email['subject'][:60]!r}] total={total_len} chars, no quoted-reply marker found (100% new)")

    print("\n--- Summary ---")
    print(f"Emails sampled: {len(emails)}")
    print(f"Emails with a detected quoted-reply marker: {marker_hits}")
    if total_chars:
        overall_pct = total_new_chars / total_chars * 100
        print(f"Overall new-text ratio: {overall_pct:.0f}% of all sampled body characters")
        print(f"(i.e. roughly {100 - overall_pct:.0f}% of what's currently being sent to Gemini is quoted/repeated history)")


if __name__ == "__main__":
    main()