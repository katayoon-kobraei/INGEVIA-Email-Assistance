import datetime
from email.header import decode_header, make_header
from email.parser import Parser
from email.utils import parseaddr
from typing import Any

import win32com.client

from src.outlook_mailbox import get_target_folder

INBOX_FOLDER_ID = 6
SENT_FOLDER_ID = 5
MAIL_ITEM_CLASS = 43
TO_RECIPIENT_TYPE = 1
CC_RECIPIENT_TYPE = 2
SMTP_ADDRESS_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
SMTP_ADDRESS_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
SENDER_SMTP_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
SENDER_SMTP_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
SENDER_NAME_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
SENDER_NAME_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
SENT_REPRESENTING_NAME_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
SENT_REPRESENTING_NAME_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
ORIGINAL_SENDER_NAME_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x005A001E"
ORIGINAL_SENDER_NAME_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x005A001F"
DISPLAY_NAME_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x3001001E"
DISPLAY_NAME_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x3001001F"
TRANSPORT_HEADERS_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
TRANSPORT_HEADERS_PROPERTY_W = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"


class _EmptyAttachments:
    """Outlook-compatible empty attachment collection."""

    Count = 0

    def Item(self, index: int) -> None:  # pragma: no cover - defensive only
        raise IndexError(index)


_EMPTY_ATTACHMENTS = _EmptyAttachments()


def _safe_get(obj: Any, property_name: str, default: Any = None) -> Any:
    """Read an Outlook COM property without aborting the whole run.

    Some Outlook folders contain meeting requests, reports, recall notices,
    partially synchronized messages, or provider-specific items.  Those items
    can raise ``<unknown>.PropertyName`` instead of returning an empty value.
    """
    try:
        value = getattr(obj, property_name)
        return default if value is None else value
    except Exception:
        return default


def _safe_call(obj: Any, method_name: str, *args: Any) -> Any:
    try:
        method = getattr(obj, method_name)
        return method(*args)
    except Exception:
        return None


def _as_utc(value: Any) -> datetime.datetime | None:
    if not isinstance(value, datetime.datetime):
        return None
    try:
        if value.tzinfo is None:
            value = value.astimezone()
        return value.astimezone(datetime.timezone.utc)
    except Exception:
        return None


def _recipient_address(recipient: Any) -> str:
    """Return the best available SMTP/display value for one recipient."""
    address_entry = _safe_get(recipient, "AddressEntry")

    exchange_user = _safe_call(address_entry, "GetExchangeUser") if address_entry else None
    smtp = str(_safe_get(exchange_user, "PrimarySmtpAddress", "") or "").strip()
    if smtp:
        return smtp

    property_accessor = _safe_get(address_entry, "PropertyAccessor") if address_entry else None
    smtp = str(_safe_call(property_accessor, "GetProperty", SMTP_ADDRESS_PROPERTY) or "").strip()
    if smtp:
        return smtp

    for obj, property_name in (
        (recipient, "Address"),
        (address_entry, "Address"),
        (recipient, "Name"),
    ):
        value = str(_safe_get(obj, property_name, "") or "").strip()
        if value:
            return value
    return ""


def _recipients_from_collection(message: Any, wanted_types: set[int]) -> str:
    recipients = _safe_get(message, "Recipients")
    count = int(_safe_get(recipients, "Count", 0) or 0)
    values: list[str] = []
    seen: set[str] = set()

    for index in range(1, count + 1):
        recipient = _safe_call(recipients, "Item", index)
        if recipient is None:
            continue
        try:
            recipient_type = int(_safe_get(recipient, "Type", 0) or 0)
        except (TypeError, ValueError):
            recipient_type = 0
        if recipient_type not in wanted_types:
            continue
        value = _recipient_address(recipient)
        normalized = value.casefold()
        if value and normalized not in seen:
            values.append(value)
            seen.add(normalized)

    return "; ".join(values)


def _recipient_text(message: Any, property_name: str, recipient_type: int) -> str:
    """Use the normal To/CC property, then fall back to Recipients."""
    value = str(_safe_get(message, property_name, "") or "").strip()
    if value:
        return value
    return _recipients_from_collection(message, {recipient_type})


def _property_accessor_value(obj: Any, *property_names: str) -> str:
    accessor = _safe_get(obj, "PropertyAccessor") if obj is not None else None
    if accessor is None:
        return ""
    for property_name in property_names:
        value = str(_safe_call(accessor, "GetProperty", property_name) or "").strip()
        if value:
            return value
    return ""


def _decode_header_text(value: str) -> str:
    value = str(value or "").strip()
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value))).strip()
    except Exception:
        return value


def _transport_sender(message: Any) -> tuple[str, str]:
    """Read the original Internet ``From`` header as a final identity fallback."""
    headers = _property_accessor_value(
        message, TRANSPORT_HEADERS_PROPERTY_W, TRANSPORT_HEADERS_PROPERTY
    )
    if not headers:
        return "", ""
    try:
        from_value = Parser().parsestr(headers).get("From", "")
    except Exception:
        return "", ""
    name, address = parseaddr(str(from_value or ""))
    return _decode_header_text(name), str(address or "").strip()


def _looks_like_display_name(value: str) -> bool:
    value = str(value or "").strip()
    if not value:
        return False
    # Exchange legacy distinguished names and bare addresses are not useful
    # human display names in the Correos sender column.
    if value.upper().startswith("/O=") or value.upper().startswith("EX:"):
        return False
    if "@" in value and " " not in value and "<" not in value:
        return False
    return True


def _sender_address(message: Any) -> str:
    """Return the sender's real SMTP address using several Outlook fallbacks."""
    sender_type = str(_safe_get(message, "SenderEmailType", "") or "").upper()
    raw_sender = str(_safe_get(message, "SenderEmailAddress", "") or "").strip()

    # Standard Internet messages normally expose the SMTP address directly.
    if sender_type != "EX" and "@" in raw_sender:
        return raw_sender

    # PidTagSenderSmtpAddress is often available even when SenderEmailAddress
    # is an Exchange legacy distinguished name.
    smtp = _property_accessor_value(
        message, SENDER_SMTP_PROPERTY_W, SENDER_SMTP_PROPERTY
    )
    if "@" in smtp:
        return smtp

    sender_object = _safe_get(message, "Sender")
    exchange_user = _safe_call(sender_object, "GetExchangeUser") if sender_object else None
    smtp = str(_safe_get(exchange_user, "PrimarySmtpAddress", "") or "").strip()
    if "@" in smtp:
        return smtp

    address_entry = _safe_get(sender_object, "AddressEntry") if sender_object else None
    smtp = _property_accessor_value(
        address_entry, SMTP_ADDRESS_PROPERTY_W, SMTP_ADDRESS_PROPERTY
    )
    if "@" in smtp:
        return smtp

    _header_name, header_address = _transport_sender(message)
    if header_address:
        return header_address

    return raw_sender


def _sender_name(message: Any) -> str:
    """Return the sender's human-readable Outlook display name.

    Outlook providers are inconsistent here: Exchange, IMAP and imported
    ``.msg`` items can expose the visible sender name through different
    properties.  Try the standard MailItem fields first, then the Sender /
    AddressEntry objects, MAPI sender-name properties and finally the original
    Internet From header.
    """
    candidates = [
        str(_safe_get(message, "SenderName", "") or "").strip(),
        str(_safe_get(message, "SentOnBehalfOfName", "") or "").strip(),
    ]

    sender_object = _safe_get(message, "Sender")
    if sender_object is not None:
        candidates.append(str(_safe_get(sender_object, "Name", "") or "").strip())

        exchange_user = _safe_call(sender_object, "GetExchangeUser")
        if exchange_user is not None:
            candidates.extend([
                str(_safe_get(exchange_user, "Name", "") or "").strip(),
                str(_safe_get(exchange_user, "FirstName", "") or "").strip() + " " +
                str(_safe_get(exchange_user, "LastName", "") or "").strip(),
            ])

        address_entry = _safe_get(sender_object, "AddressEntry")
        if address_entry is not None:
            candidates.append(str(_safe_get(address_entry, "Name", "") or "").strip())
            candidates.append(
                _property_accessor_value(
                    address_entry, DISPLAY_NAME_PROPERTY_W, DISPLAY_NAME_PROPERTY
                )
            )

    candidates.extend([
        _property_accessor_value(message, SENDER_NAME_PROPERTY_W, SENDER_NAME_PROPERTY),
        _property_accessor_value(
            message, SENT_REPRESENTING_NAME_PROPERTY_W, SENT_REPRESENTING_NAME_PROPERTY
        ),
        _property_accessor_value(
            message, ORIGINAL_SENDER_NAME_PROPERTY_W, ORIGINAL_SENDER_NAME_PROPERTY
        ),
    ])

    header_name, _header_address = _transport_sender(message)
    candidates.append(header_name)

    for value in candidates:
        value = " ".join(_decode_header_text(value).split())
        if _looks_like_display_name(value):
            # If Outlook gives a complete RFC-style mailbox string here, keep
            # only its display-name part. The email is rendered separately.
            parsed_name, parsed_email = parseaddr(value)
            if parsed_email and parsed_name:
                value = _decode_header_text(parsed_name)
            if value:
                return value

    return ""


def resolve_sender_identity(message: Any) -> tuple[str, str]:
    """Return ``(display_name, smtp_address)`` for a MailItem."""
    return _sender_name(message), _sender_address(message)


def _fetch_from_folder(outlook: Any, folder_id: int, minutes_back: float, direction: str) -> list[dict]:
    folder = get_target_folder(outlook, folder_id)
    messages = folder.Items
    store_id = str(_safe_get(folder, "StoreID", "") or "")
    time_field = "ReceivedTime" if direction == "ENTRANTE" else "SentOn"

    sorted_descending = True
    try:
        messages.Sort(f"[{time_field}]", True)
    except Exception:
        # If a provider rejects Sort(), scan all items rather than assuming
        # their existing order is chronological.
        sorted_descending = False

    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=minutes_back)
    results: list[dict] = []
    skipped_unsupported = 0
    skipped_unreadable = 0

    item_count = int(_safe_get(messages, "Count", 0) or 0)
    for index in range(1, item_count + 1):
        message = _safe_call(messages, "Item", index)
        if message is None:
            skipped_unreadable += 1
            continue

        item_class = _safe_get(message, "Class")
        if item_class is not None:
            try:
                if int(item_class) != MAIL_ITEM_CLASS:
                    skipped_unsupported += 1
                    continue
            except (TypeError, ValueError):
                skipped_unsupported += 1
                continue

        timestamp = _safe_get(message, time_field)
        timestamp_utc = _as_utc(timestamp)
        if timestamp_utc is None:
            skipped_unreadable += 1
            continue
        if timestamp_utc < cutoff:
            if sorted_descending:
                break
            continue

        entry_id = str(_safe_get(message, "EntryID", "") or "").strip()
        if not entry_id:
            skipped_unreadable += 1
            continue

        was_unread = bool(_safe_get(message, "UnRead", False))
        to_text = _recipient_text(message, "To", TO_RECIPIENT_TYPE)
        cc_text = _recipient_text(message, "CC", CC_RECIPIENT_TYPE)
        body = str(_safe_get(message, "Body", "") or "")

        actual_sender_email = _sender_address(message)
        actual_sender_name = _sender_name(message)

        results.append(
            {
                "id": entry_id,
                "store_id": store_id,
                "subject": str(_safe_get(message, "Subject", "(Sin asunto)") or "(Sin asunto)"),
                # Keep the existing routing semantics unchanged: ``sender`` is
                # populated only for incoming messages.  The two fields below
                # are dedicated display/audit fields and are safe for both
                # incoming and outgoing messages.
                "sender": actual_sender_email if direction == "ENTRANTE" else None,
                "sender_name": actual_sender_name,
                "sender_email": actual_sender_email,
                "recipient": to_text if direction == "SALIENTE" else None,
                "to": to_text,
                "cc": cc_text,
                "timestamp": timestamp,
                "body": body,
                "attachments": _safe_get(message, "Attachments", _EMPTY_ATTACHMENTS),
                "direction": direction,
            }
        )

        # Reading Body can mark an IMAP message as read. Restore the previous
        # unread state without allowing a write failure to stop processing.
        if was_unread and not bool(_safe_get(message, "UnRead", was_unread)):
            try:
                message.UnRead = True
                message.Save()
            except Exception:
                pass

    skipped_total = skipped_unsupported + skipped_unreadable
    if skipped_total:
        print(
            f"Outlook: skipped {skipped_total} unsupported/unreadable item(s) "
            f"in {'Inbox' if direction == 'ENTRANTE' else 'Sent Items'}; "
            "normal emails will continue to be processed."
        )

    return results


def get_recent_emails(minutes_back: int = 30, outlook: Any = None) -> list[dict]:
    """Return recent standard mail items from Inbox and Sent Items.

    ``outlook`` may be an already-open MAPI namespace reused by the pipeline.
    """
    if outlook is None:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbound = _fetch_from_folder(outlook, INBOX_FOLDER_ID, minutes_back, "ENTRANTE")
    outbound = _fetch_from_folder(outlook, SENT_FOLDER_ID, minutes_back, "SALIENTE")
    return inbound + outbound
