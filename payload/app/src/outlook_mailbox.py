from __future__ import annotations

from typing import Any

from src.config import TARGET_MAILBOX


def _safe_get(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
        return default if value is None else value
    except Exception:
        return default


def _normalized(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def get_target_folder(namespace: Any, folder_id: int, target_mailbox: str | None = None) -> Any:
    """Return a default folder from the explicitly configured mailbox.

    The function deliberately fails instead of silently processing another
    Windows user's default mailbox when the configured boss mailbox is not
    accessible.
    """
    target = (target_mailbox or TARGET_MAILBOX or "").strip()
    target_key = _normalized(target)
    if not target_key:
        return namespace.GetDefaultFolder(folder_id)

    recipient = None
    resolved_name = ""
    try:
        recipient = namespace.CreateRecipient(target)
        recipient.Resolve()
        if bool(_safe_get(recipient, "Resolved", False)):
            resolved_name = _normalized(_safe_get(recipient, "Name", ""))
    except Exception:
        recipient = None

    # Primary/account mailboxes exposed in the current Outlook profile.
    accounts = _safe_get(namespace, "Accounts")
    account_count = int(_safe_get(accounts, "Count", 0) or 0)
    for index in range(1, account_count + 1):
        try:
            account = accounts.Item(index)
            smtp = _normalized(_safe_get(account, "SmtpAddress", ""))
            display_name = _normalized(_safe_get(account, "DisplayName", ""))
            if target_key in {smtp, display_name} or resolved_name == display_name or (
                display_name and (target_key in display_name or display_name in target_key)
            ):
                store = _safe_get(account, "DeliveryStore")
                if store is not None:
                    return store.GetDefaultFolder(folder_id)
        except Exception:
            continue

    # Additional/shared stores mounted in Outlook's folder tree.
    stores = _safe_get(namespace, "Stores")
    store_count = int(_safe_get(stores, "Count", 0) or 0)
    for index in range(1, store_count + 1):
        try:
            store = stores.Item(index)
            display_name = _normalized(_safe_get(store, "DisplayName", ""))
            if target_key == display_name or resolved_name == display_name or (
                display_name and (target_key in display_name or display_name in target_key)
            ):
                return store.GetDefaultFolder(folder_id)
        except Exception:
            continue

    # Delegated/shared Exchange folder resolution by SMTP address. Outlook
    # does not expose every default folder through this API (notably Sent Mail),
    # so a fully mounted Store is preferred above.
    if recipient is not None and bool(_safe_get(recipient, "Resolved", False)):
        try:
            return namespace.GetSharedDefaultFolder(recipient, folder_id)
        except Exception:
            pass

    raise RuntimeError(
        f"Outlook cannot access the configured mailbox '{target}'. "
        "Add that mailbox to the Outlook profile or ask the Microsoft 365 "
        "administrator to grant the required mailbox permissions."
    )


def target_mailbox_label() -> str:
    return TARGET_MAILBOX or "default Outlook mailbox"
