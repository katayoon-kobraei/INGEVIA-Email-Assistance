"""
run_state.py

Tracks the timestamp of the last successful pipeline run, so
pipeline.py can compute exactly how far back to look for new emails
instead of relying on a fixed window (LOOKBACK_MINUTES).

This matters now that the scheduled task only runs during business
hours (05:30-20:00): the first run of the day has to cover the entire
overnight gap since the previous evening's last run (roughly 9.5
hours), which is much longer than the normal ~30-minute gap between
runs during the day. Storing the actual last-run time -- rather than
hardcoding "the first run covers N hours overnight" -- means this
self-corrects for weekends, holidays, or any run that gets skipped for
an unrelated reason (PC off, Outlook down, scheduler paused): the next
run just closes whatever gap actually happened, however long it was.

Safe to look back further than strictly needed -- pipeline.py already
skips any email whose ID is in _processed_ids.json, so a wider window
never causes double-filing. The only cost of a wider window is a
slightly slower Outlook scan.

Where this goes: src/output/run_state.py (same folder as
flag_state.py, pending_copy_state.py).
"""

import datetime
import json
import os

STATE_FILENAME = "_last_run.json"


def _state_path(output_root):
    return os.path.join(output_root, STATE_FILENAME)


def load_last_run_time(output_root):
    """Returns the UTC datetime of the last successful run, or None if
    this is the first run ever (or the state file is missing/corrupt --
    treated the same as "no prior run" rather than crashing the
    pipeline over a state-tracking file)."""
    path = _state_path(output_root)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        when = datetime.datetime.fromisoformat(data["last_run_utc"])
        if when.tzinfo is None:
            when = when.replace(tzinfo=datetime.timezone.utc)
        return when
    except Exception:
        return None


def save_last_run_time(output_root, when_utc):
    """Overwrites the stored last-run timestamp. Never raises -- a
    failure to persist this is not worth breaking the pipeline over;
    worst case the next run just falls back to LOOKBACK_MINUTES instead
    of the exact gap."""
    path = _state_path(output_root)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"last_run_utc": when_utc.isoformat()}, f)
    except Exception as e:
        print(f"Could not save last-run timestamp: {e}")