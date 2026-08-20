from __future__ import annotations

import os
from contextlib import contextmanager
from pathlib import Path


class ProcessingAlreadyRunning(RuntimeError):
    pass


def _lock_path() -> Path:
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        root = Path(local_appdata) / "INGEVIA Email Assistant"
    else:
        root = Path(__file__).resolve().parent.parent
    root.mkdir(parents=True, exist_ok=True)
    return root / "email_processing.lock"


@contextmanager
def processing_lock():
    """Cross-process lock for one Email Assistant processing run per PC.

    On Windows the byte-range lock is owned by the process and is released
    automatically if that process exits or crashes, so stale lock files are
    harmless.  The file itself may remain on disk; only the OS lock matters.
    """
    path = _lock_path()
    handle = path.open("a+b")
    acquired = False
    try:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"0")
            handle.flush()
        handle.seek(0)

        if os.name == "nt":
            import msvcrt

            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                acquired = True
            except OSError as exc:
                raise ProcessingAlreadyRunning(
                    "Email processing is already running on this computer."
                ) from exc
        else:
            # Development/test fallback. Production deployment is Windows.
            import fcntl

            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
            except OSError as exc:
                raise ProcessingAlreadyRunning(
                    "Email processing is already running on this computer."
                ) from exc

        yield path
    finally:
        if acquired:
            try:
                handle.seek(0)
                if os.name == "nt":
                    import msvcrt

                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
        handle.close()
