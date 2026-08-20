from __future__ import annotations

import datetime as dt
import json
import os
from pathlib import Path

from src.config import OUTPUT_ROOT


STATE_FILE = Path(OUTPUT_ROOT) / "_gemini_circuit_state.json"


def _read() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write(data: dict) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp = STATE_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, STATE_FILE)


def open_circuit(minutes: int, reason: str) -> dt.datetime:
    until = dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=max(1, int(minutes)))
    _write({"open_until": until.isoformat(), "reason": str(reason)[:1000]})
    return until


def circuit_status() -> tuple[bool, dt.datetime | None, str]:
    data = _read()
    text = str(data.get("open_until") or "")
    if not text:
        return False, None, ""
    try:
        until = dt.datetime.fromisoformat(text)
        if until.tzinfo is None:
            until = until.replace(tzinfo=dt.timezone.utc)
    except ValueError:
        return False, None, ""
    if until > dt.datetime.now(dt.timezone.utc):
        return True, until, str(data.get("reason") or "")
    return False, until, str(data.get("reason") or "")


def clear_circuit() -> None:
    try:
        STATE_FILE.unlink(missing_ok=True)
    except OSError:
        pass
