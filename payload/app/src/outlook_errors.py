from __future__ import annotations

_RESOURCE_MARKERS = (
    "outlook has exhausted all shared resources",
    "outlook ha agotado todos los recursos compartidos",
    "cannot open the outlook window",
    "no se puede abrir la ventana de outlook",
    "cannot open the set of folders",
    "no se puede abrir el conjunto de carpetas",
    "mapi_e_not_enough_resources",
    "not enough memory or system resources",
    "insufficient memory",
    "0x8004010d",
    "0x8007000e",
    # COM startup failures observed when classic Outlook is already in a bad
    # MAPI/resource state. Treat them as transient resource failures so the
    # supervisor performs the same clean-worker retry instead of a generic exit 1.
    "operation unavailable",
    "operación no disponible",
    "server execution failed",
    "error en la ejecución de servidor",
    "-2147221021",
    "-2146959355",
)


def is_outlook_resource_error(exc: BaseException) -> bool:
    parts = [str(exc), repr(exc)]
    for arg in getattr(exc, "args", ()):
        parts.append(str(arg))
    text = " ".join(parts).casefold()
    return any(marker in text for marker in _RESOURCE_MARKERS)


_MAILBOX_MARKERS = (
    "outlook cannot access the configured mailbox",
    "cannot access the configured mailbox",
)


def is_outlook_mailbox_unavailable(exc: BaseException) -> bool:
    parts = [str(exc), repr(exc)]
    for arg in getattr(exc, "args", ()):
        parts.append(str(arg))
    text = " ".join(parts).casefold()
    return any(marker in text for marker in _MAILBOX_MARKERS)
