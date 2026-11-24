"""Activity journal of the Outlook pass, read by the Dashboard.

Appends one JSON event per line (JSONL, UTF-8), written only by the pass.
Home-server convention (home-server docs/architecture.md, System tab):
``{"time", "action", "object"}`` plus, when not empty, ``"from"``, ``"to"``,
``"detail"``, ``"engine"``, ``"model"``, ``"confidence"`` and ``"run"``.
``object``, ``from`` and ``to`` are cut at 300 characters, ``detail`` at 1000.
Texts are French (read by the owner) and may hold subjects and addresses:
the file stays private (mode 600) and never feeds a notification.
Above 5 MB the file is rewritten atomically with the events of the last 90 days,
20,000 lines at most. Writing never raises: a journal failure must not stop a pass.
Path: ``$OUTLOOK_SORTER_ACTIVITY``, defaulting to ``activity.jsonl`` next to
``senders.json`` (in ``$OUTLOOK_SORTER_DATA``, or this folder).
"""

import datetime
import json
import os
from pathlib import Path
from typing import Any

PATH = Path(
    os.environ.get("OUTLOOK_SORTER_ACTIVITY")
    or Path(os.environ.get("OUTLOOK_SORTER_DATA") or Path(__file__).resolve().parent) / "activity.jsonl"
)
MAX_BYTES = 5_000_000
KEEP_DAYS = 90
KEEP_LINES = 20_000
LIMITS = {"object": 300, "from": 300, "to": 300, "detail": 1000}
# Rotate at most once per run: 20,000 long lines can remain above 5 MB.
rotated = False


def now() -> str:
    """Return the current local ISO-8601 timestamp with timezone offset."""
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def _encode_event(action: str, obj: str, fields: dict[str, Any]) -> str:
    """Serialize one journal event to a single JSONL line with safe line breaks."""
    event: dict[str, Any] = {"time": now(), "action": action, "object": obj}
    event.update({k: v for k, v in fields.items() if v not in (None, "", [])})
    for key, limit in LIMITS.items():
        if isinstance(event.get(key), str):
            event[key] = event[key][:limit]
    line = json.dumps(event, ensure_ascii=False)
    # Escape Unicode line breaks so line-oriented readers never split one event.
    for char in "\x85\u2028\u2029":
        line = line.replace(char, f"\\u{ord(char):04x}")
    return line


def log(action: str, obj: str, **fields: Any) -> None:
    """Append one event to the activity journal; empty fields are omitted."""
    global rotated
    try:
        line = _encode_event(action, obj, fields)
        fd = os.open(PATH, os.O_WRONLY | os.O_APPEND | os.O_CREAT, 0o600)
        try:
            os.write(fd, (line + "\n").encode("utf-8", "replace"))
            size = os.fstat(fd).st_size
        finally:
            os.close(fd)
        if size > MAX_BYTES and not rotated:
            rotated = True
            rotate()
    except OSError, TypeError, ValueError:
        pass


def rotate() -> None:
    """Keep events from the last 90 days (at most 20,000 lines), dropping unreadable lines."""
    since = datetime.datetime.now().astimezone() - datetime.timedelta(days=KEEP_DAYS)
    kept: list[str] = []
    with open(PATH, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            try:
                if datetime.datetime.fromisoformat(json.loads(line)["time"]) >= since:
                    kept.append(line.rstrip("\n") + "\n")
            except ValueError, KeyError, TypeError:
                continue
    tmp = PATH.with_name(f".{PATH.name}.tmp")
    try:
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.writelines(kept[-KEEP_LINES:])
        os.replace(tmp, PATH)
    finally:
        tmp.unlink(missing_ok=True)
