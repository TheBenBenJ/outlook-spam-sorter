"""Activity journal of the Outlook pass, read by the Dashboard: one JSON event per line, appended (JSONL, UTF-8), written only by the pass.

Convention of the home server (home-server docs/architecture.md, System tab): {"time", "action", "object"} plus, when not empty, "from",
"to", "detail", "engine", "model", "confidence" and "run". object, from and to are cut at 300 characters, detail at 1000. The texts are
French (Benjamin reads them) and may hold subjects and addresses: the file stays private (mode 600) and never feeds a notification.
Above 5 MB the file is rewritten (a temporary file in the same folder, then a rename) with the events of the last 90 days, 20,000 lines
at most. Writing never raises: a journal that cannot be written must not stop the pass.
Path: $OUTLOOK_SORTER_ACTIVITY, by default activity.jsonl next to the pass's state (senders.json, outlook-state.json).
"""
import datetime, json, os
from pathlib import Path

PATH = Path(os.environ.get("OUTLOOK_SORTER_ACTIVITY") or Path(__file__).resolve().parent / "activity.jsonl")
MAX_BYTES, KEEP_DAYS, KEEP_LINES = 5_000_000, 90, 20_000
LIMITS = {"object": 300, "from": 300, "to": 300, "detail": 1000}
rotated = False   # once per run: 20,000 long lines may stay above 5 MB, and rewriting the file at every event would be wasteful


def now():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def log(action, obj, **fields):
    """Appends one event; empty fields are left out."""
    global rotated
    try:
        ev = {"time": now(), "action": action, "object": obj}
        ev.update({k: v for k, v in fields.items() if v not in (None, "", [])})
        for k, n in LIMITS.items():
            if isinstance(ev.get(k), str):
                ev[k] = ev[k][:n]
        line = json.dumps(ev, ensure_ascii=False)
        for c in "\x85  ":   # line breaks for some readers (str.splitlines): escaped, the JSON value stays the same
            line = line.replace(c, f"\\u{ord(c):04x}")
        fd = os.open(PATH, os.O_WRONLY | os.O_APPEND | os.O_CREAT, 0o600)
        try:
            os.write(fd, (line + "\n").encode("utf-8", "replace"))
            size = os.fstat(fd).st_size
        finally:
            os.close(fd)
        if size > MAX_BYTES and not rotated:
            rotated = True
            rotate()
    except Exception:
        pass


def rotate():
    """Keeps the events of the last 90 days, the last 20,000 at most. An unreadable line is dropped."""
    since = datetime.datetime.now().astimezone() - datetime.timedelta(days=KEEP_DAYS)
    kept = []
    with open(PATH, encoding="utf-8", errors="replace") as f:
        for line in f:
            try:
                if datetime.datetime.fromisoformat(json.loads(line)["time"]) >= since:
                    kept.append(line.rstrip("\n") + "\n")
            except (ValueError, KeyError, TypeError):
                continue
    tmp = PATH.with_name(f".{PATH.name}.tmp")
    try:
        with os.fdopen(os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600), "w", encoding="utf-8") as f:
            f.writelines(kept[-KEEP_LINES:])
        os.replace(tmp, PATH)
    finally:
        tmp.unlink(missing_ok=True)
