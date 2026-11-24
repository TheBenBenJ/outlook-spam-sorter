#!/usr/bin/env python3
"""Daily pass over Outlook.com by IMAP (XOAUTH2), without a browser.

Mail content is hostile: actions never depend on what a mail says. The pass opens no
link, sends nothing, never sets deletion flags, and never expunges (deleting moves
messages to Deleted Items, kept 30 days).

Usage:
    imap-pass.py [trial]   trial = dry run (read-only folders, nothing moved/recorded)

Junk folder: known senders go back to the inbox, confirmed spam goes to Deleted Items,
the rest waits for a person (analyze.py, classify.py).
Inbox: stuck mails from trusted senders are deposited into Gmail via ``reinjecter``,
then archived; duplicates, noise, and old unknown mails are archived.
Optional ``$OUTLOOK_SORTER_NOTIFY`` is invoked as ``<cmd> "Tri mails Outlook" "<text>"``
on failure or when the count of mails to review rises (never carrying addresses/subjects).
State lives in ``$OUTLOOK_SORTER_DATA`` (default: this folder).
"""

import email
import email.errors
import email.header
import email.utils
import imaplib
import importlib
import json
import os
import re
import signal
import subprocess
import sys
import time
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("OUTLOOK_SORTER_DATA") or HERE)
WORK = DATA / "work"
EML = WORK / "eml"
SPAM_CONFIDENCE = 0.6
MAX_DURATION = 30 * 60
CHUNK = 200
# Journal locations and verdicts as read by the owner (not raw IMAP folder names).
PLACE = {
    "junk": "Courrier indésirable",
    "inbox": "Boîte de réception",
    "outlook": "Boîte de réception Outlook",
    "trash": "Éléments supprimés",
    "archive": "Archivage",
    "gmail": "Gmail",
}
WORDS = {"legitime": "légitime", "spam": "spam", "douteux": "douteux"}


def _import_sibling(name: str) -> Any:
    """Import a sibling module under `python3 -I`, ensuring HERE precedes other paths."""
    sys.path.insert(0, str(HERE))
    return importlib.import_module(name)


def _parse_dry_run(argv: list[str]) -> bool:
    """Validate CLI arguments and return True when running in dry-run ``trial`` mode."""
    if argv not in ([], ["trial"]):
        sys.exit("usage: imap-pass.py [trial]")
    return argv == ["trial"]


activity = _import_sibling("activity")
outlook_imap = _import_sibling("outlook_imap")
DRY_RUN = _parse_dry_run(sys.argv[1:])
RUN = activity.now()
# Mails already reported as "à vérifier" (folder:UID:verdict), avoided on repeat passes.
WAITING = activity.PATH.with_name("activity-waiting.json")
reported: set[str] = set()
waiting_now: set[str] = set()


@dataclass(slots=True)
class InboxBatch:
    """Context shared while depositing and archiving inbox rows for one pass."""

    client: Any
    verdicts: dict[str, dict[str, Any]]
    subjects: dict[str, str]
    script: Any
    report: dict[str, Any]


def log(*parts: Any) -> None:
    """Print a timestamped log line to stdout."""
    print(time.strftime("%H:%M:%S"), *parts, flush=True)


def event(
    action: str,
    obj: str,
    src: str | None = None,
    dst: str | None = None,
    detail: str | None = None,
    v: dict[str, Any] | None = None,
) -> None:
    """Append one journal event unless running in dry-run mode."""
    if DRY_RUN:
        return
    verdict = v or {}
    raw_conf = verdict.get("confiance") if verdict.get("verdict") else None
    confidence = min(1.0, max(0.0, float(raw_conf))) if isinstance(raw_conf, (int, float)) else None
    activity.log(
        action,
        obj or "(sans objet)",
        **{"from": src, "to": dst},
        detail=detail,
        engine=verdict.get("engine"),
        model=verdict.get("model"),
        confidence=confidence,
        run=RUN,
    )


def judged(verdict: dict[str, Any]) -> str:
    """Format a model verdict and clamped confidence percentage in French."""
    label = WORDS.get(verdict.get("verdict", ""), verdict.get("verdict"))
    pct = round(100 * min(1.0, max(0.0, float(verdict.get("confiance", 0)))))
    return f"{label} (sûr à {pct} %)"


def detail(row: dict[str, Any], verdict: dict[str, Any], note: str = "") -> str:
    """Build the French detail text explaining why a mail received its action."""
    text = f"{row['from'] or '(sans adresse)'} : {row.get('motif', '')}"
    if verdict.get("verdict"):
        text += f", jugé {judged(verdict)}"
    return f"{text} ; {note}" if note else text


def review_once(
    key: str,
    obj: str,
    src: str,
    text: str,
    verdict: dict[str, Any] | None = None,
) -> None:
    """Emit an ``à vérifier`` event at most once per ``folder:UID:verdict`` key."""
    waiting_now.add(key)
    if key not in reported:
        event("à vérifier", obj, src, None, text, verdict)


def clean(value: Any, limit: int = 300) -> str:
    """Strip control/bidi characters, collapse whitespace, and truncate."""
    cleaned = re.sub(r"[\x00-\x1f\x7f\u200b-\u200f\u202a-\u202e\u2066-\u2069]", " ", str(value or ""))
    return re.sub(r"\s+", " ", cleaned).strip()[:limit]


def run_script(name: str, script_args: Any = (), stdin: str | None = None) -> str:
    """Run a sibling Python helper in isolated mode and return its stdout."""
    result = subprocess.run(
        [sys.executable, "-I", str(HERE / name), *map(str, script_args)],
        input=stdin,
        capture_output=True,
        text=True,
        timeout=900,
    )
    if result.returncode != 0:
        raise RuntimeError(f"{name} : {(result.stderr or result.stdout).strip()[-200:]}")
    return result.stdout


def analyze(filename: str, snap: list[dict[str, Any]], analyze_args: list[str]) -> dict[str, Any]:
    """Write a snapshot into work/ and parse the JSON output of analyze.py."""
    (WORK / filename).write_text(json.dumps(snap, ensure_ascii=False))
    return json.loads(run_script("analyze.py", [*analyze_args, filename]))


def judge(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Run classify.py for rows and return verdicts keyed by mail UID.

    Logs an ``échec`` journal event if classify.py fails or leaves rows without verdicts.
    """
    if not rows:
        return {}
    payload = json.dumps([{"n": idx, "from": row["from"], "objet": row["objet"]} for idx, row in enumerate(rows)])
    try:
        result = subprocess.run(
            [sys.executable, "-I", str(HERE / "classify.py")],
            input=payload,
            capture_output=True,
            text=True,
            timeout=900,
        )
        for line in result.stderr.splitlines():
            log(" ", clean(line, 240))
        if result.returncode != 0:
            raise RuntimeError(f"classify.py : {result.stderr.strip()[-200:]}")
        raw_out = json.loads(result.stdout)
        verdicts = {row["id"]: raw_out.get(str(idx), {}) for idx, row in enumerate(rows)}
    except (OSError, subprocess.SubprocessError, ValueError, RuntimeError) as error:
        log("model verdicts unavailable:", clean(error, 200))
        event(
            "échec",
            "passage Outlook",
            detail=(
                "avis du modèle local indisponibles, les mails à juger attendent le prochain passage : "
                f"{clean(error, 800)}"
            ),
        )
        return {}
    missing = sum(not v.get("verdict") for v in verdicts.values())
    if missing:
        why = (result.stderr.strip().splitlines() or ["sans raison donnée"])[-1]
        plural = "s" if missing > 1 else ""
        verb = "ils attendent" if missing > 1 else "il attend"
        event(
            "échec",
            "passage Outlook",
            detail=(
                f"{missing} mail{plural} sur {len(rows)} sans avis du modèle local, "
                f"{verb} le prochain passage : {clean(why, 800)}"
            ),
        )
    return verdicts


def header(msg: email.message.Message, name: str) -> str:
    """Decode an RFC 2047 email header, falling back to the raw value on malformed input."""
    raw = msg.get(name, "")
    try:
        return str(email.header.make_header(email.header.decode_header(raw)))
    except ValueError, LookupError, TypeError, email.errors.MessageError:
        return raw


def check(typ: str, data: Any, what: str) -> Any:
    """Ensure an IMAP command returned ``OK``, otherwise raise a descriptive RuntimeError."""
    if typ != "OK":
        raise RuntimeError(f"{what} refusé par Outlook : {clean(data, 120)}")
    return data


def snapshot(client: Any, folder: str) -> list[dict[str, Any]]:
    """Fetch mail headers and received dates from an IMAP folder as analyze.py expects."""
    check(*client.select(f'"{folder}"', readonly=DRY_RUN), f"ouverture de {folder}")
    uids = check(*client.uid("SEARCH", None, "ALL"), "recherche")[0].split()
    rows: list[dict[str, Any]] = []
    for index in range(0, len(uids), CHUNK):
        data = check(
            *client.uid(
                "FETCH",
                b",".join(uids[index : index + CHUNK]).decode(),
                "(UID INTERNALDATE BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])",
            ),
            "lecture des en-têtes",
        )
        for part in data:
            if not isinstance(part, tuple):
                continue
            uid = re.search(rb"UID (\d+)", part[0])
            if not uid:
                continue
            when = imaplib.Internaldate2tuple(part[0])
            # Spam headers frequently contain raw UTF-8 bytes outside RFC 2047 encoding.
            msg = email.message_from_string(part[1].decode("utf-8", "replace"))
            name, address = email.utils.parseaddr(header(msg, "From"))
            subj = header(msg, "Subject")
            rows.append(
                {
                    "id": uid[1].decode(),
                    "from": address.strip().lower(),
                    "l": clean(f"{name} {subj}"),
                    "date": time.strftime("%Y-%m-%d", when) if when else None,
                    "subject": clean(subj),
                }
            )
    return rows


def move(
    client: Any,
    rows: list[dict[str, Any]],
    dest: str,
    moved: Callable[[dict[str, Any]], None] | None = None,
) -> int:
    """Move rows by UID in chunks (never STORE/EXPUNGE) and invoke ``moved`` per row."""
    done = 0
    for index in range(0, len(rows), CHUNK):
        chunk = rows[index : index + CHUNK]
        check(*client.uid("MOVE", ",".join(row["id"] for row in chunk), f'"{dest}"'), "déplacement")
        done += len(chunk)
        if moved is not None:
            for row in chunk:
                moved(row)
    return done


def _junk_review_note(
    row: dict[str, Any],
    verdict_name: str | None,
    by_class: dict[str, list[dict[str, Any]]],
) -> str:
    """Explain in French why a junk mail was left in place for human review."""
    if not verdict_name:
        return "sans avis du modèle, nouvel essai au prochain passage"
    if verdict_name == "spam" and row in by_class["connu_gmail"]:
        return "un expéditeur connu n'est jamais supprimé"
    if verdict_name == "spam":
        return f"pas assez sûr pour supprimer (il faut {round(100 * SPAM_CONFIDENCE)} %)"
    if verdict_name == "legitime" and row in by_class["a_classer"]:
        return "seul un expéditeur connu est remis en boîte"
    return ""


def _log_dry_run_junk(restore: list[dict[str, Any]], remove: list[dict[str, Any]]) -> None:
    """Log the junk actions that would have run outside dry-run mode."""
    for row in restore:
        log("  to restore:", row["from"], "|", row["objet"][:70])
    for row in remove:
        log("  spam:", row["from"], "|", row["objet"][:70])


def junk_folder(client: Any, names: dict[str, str]) -> dict[str, Any]:
    """Classify the junk folder, restore legitimate senders, move sure spam to Deleted Items."""
    snap = snapshot(client, names["junk"])
    subjects = {row["id"]: row["subject"] for row in snap}
    by_class = analyze("junk.json", snap, ["--lines"])["lignes"]
    to_judge = by_class["connu_gmail"] + by_class["a_classer"]
    verdicts = judge(to_judge)

    def verdict_of(row: dict[str, Any]) -> dict[str, Any]:
        return verdicts.get(row["id"], {})

    restore = by_class["legitime"] + [
        row for row in by_class["connu_gmail"] if verdict_of(row).get("verdict") == "legitime"
    ]
    # Strict filtering: unknown senders judged spam (>= 60%) or doubtful go to Deleted Items (30 days).
    new_spam = [
        row
        for row in by_class["a_classer"]
        if (verdict_of(row).get("verdict") == "spam" and verdict_of(row).get("confiance", 0) >= SPAM_CONFIDENCE)
        or verdict_of(row).get("verdict") == "douteux"
    ]
    remove = by_class["spam_connu"] + new_spam
    waiting = [row for row in to_judge if row not in restore and row not in remove]
    log(f"junk: {len(snap)} rows, {len(restore)} to restore, {len(remove)} spam, {len(waiting)} to review")
    for row in waiting:
        log("  to review:", row["from"], "|", row["objet"][:70], json.dumps(verdict_of(row)))
    report: dict[str, Any] = {
        "examined": len(snap),
        "restored": [],
        "deleted": 0,
        "waiting": len(waiting),
        "new_spam": new_spam,
    }
    if DRY_RUN:
        _log_dry_run_junk(restore, remove)
        return report

    # Restore to INBOX so the subsequent inbox sweep recovers them into Gmail.
    move(
        client,
        restore,
        "INBOX",
        lambda row: event(
            "remis en boîte",
            subjects[row["id"]],
            PLACE["junk"],
            PLACE["inbox"],
            detail(row, verdict_of(row)),
            verdict_of(row),
        ),
    )
    report["restored"] = [row["from"] for row in restore]
    report["deleted"] = move(
        client,
        remove,
        names["trash"],
        lambda row: event(
            "spam supprimé",
            subjects[row["id"]],
            PLACE["junk"],
            PLACE["trash"],
            detail(row, verdict_of(row)),
            verdict_of(row),
        ),
    )
    for row in waiting:
        verdict = verdict_of(row)
        verdict_name = verdict.get("verdict")
        note = _junk_review_note(row, verdict_name, by_class)
        review_once(
            f"{names['junk']}:{row['id']}:{verdict_name or ''}",
            subjects[row["id"]],
            PLACE["junk"],
            detail(row, verdict, note),
            verdict,
        )
    log(f"junk: {len(restore)} restored, {report['deleted']} deleted")
    return report


def deposit(client: Any, uid: str, script: Any) -> dict[str, Any]:
    """Fetch a raw .eml with BODY.PEEK[] and hand it to the reinjection script."""
    data = check(*client.uid("FETCH", uid, "(BODY.PEEK[])"), "lecture du mail")
    raw = next((part[1] for part in data if isinstance(part, tuple)), None)
    if not raw:
        return {"statut": "echec", "motif": "mail introuvable"}
    eml_path = EML / f"{uid}.eml"
    eml_path.write_bytes(raw)
    cmd = script if isinstance(script, list) else [script]
    result = subprocess.run([*cmd, str(eml_path)], capture_output=True, text=True, timeout=300)
    last = (result.stdout.strip().splitlines() or [""])[-1]
    try:
        return json.loads(last)
    except ValueError:
        return {"statut": "echec", "motif": clean(result.stderr, 120) or "pas de réponse"}
    finally:
        eml_path.unlink(missing_ok=True)


def _deposit_candidates(batch: InboxBatch, to_recover: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deposit trusted inbox rows into Gmail and return those ready to archive."""
    done: list[dict[str, Any]] = []
    for row in to_recover:
        reply = deposit(batch.client, row["id"], batch.script)
        verdict = batch.verdicts.get(row["id"], {})
        status = reply.get("statut")
        if status == "refuse":
            batch.report["waiting"] += 1
            log("  refused:", reply.get("de"), "|", clean(reply.get("objet"), 50), reply.get("motif"))
            review_once(
                f"INBOX:{row['id']}:refuse",
                batch.subjects[row["id"]],
                PLACE["outlook"],
                detail(row, {}, f"dépôt dans Gmail refusé : {clean(reply.get('motif'), 300)}"),
            )
        elif status in ("depose", "deja"):
            batch.report["recovered"].add(row["from"])
            batch.report["deposited"] += status == "depose"
            if status == "depose":
                event(
                    "récupéré dans Gmail",
                    batch.subjects[row["id"]],
                    PLACE["outlook"],
                    PLACE["gmail"],
                    detail(row, verdict),
                    verdict,
                )
            motif = "récupéré dans Gmail" if status == "depose" else "Gmail l'avait déjà"
            done.append({**row, "motif": motif})
        else:
            batch.report["remaining"] += 1
            log("  deposit failed:", reply.get("de"), reply.get("motif"))
            reason = clean(reply.get("motif"), 300) or "sans raison donnée"
            event(
                "échec",
                batch.subjects[row["id"]],
                PLACE["outlook"],
                PLACE["gmail"],
                detail(row, {}, f"dépôt dans Gmail échoué : {reason}, nouvel essai au prochain passage"),
            )
    return done


def _log_dry_run_inbox(to_recover: list[dict[str, Any]], to_archive: list[dict[str, Any]]) -> None:
    """Log inbox recovery and archive candidates during a dry run."""
    for row in to_recover:
        log("  to recover:", row["from"], "|", row["objet"][:70])
    for row in to_archive:
        log("  to archive:", row["from"], "|", row["objet"][:70])


def inbox(client: Any, names: dict[str, str], script: Any) -> dict[str, Any]:
    """Recover stuck trusted inbox messages into Gmail and archive processed or old rows."""
    snap = snapshot(client, "INBOX")
    subjects = {row["id"]: row["subject"] for row in snap}
    analysis = analyze("inbox.json", snap, ["inbox"])
    verdicts = judge(analysis["recuperer"])
    to_recover = [row for row in analysis["recuperer"] if verdicts.get(row["id"], {}).get("verdict") == "legitime"]
    waiting = [row for row in analysis["recuperer"] if row not in to_recover]
    review_count = analysis["a_verifier_total"] + len(waiting)
    remaining_count = analysis["recuperer_restants"] + analysis["archiver_restants"]
    log(
        f"inbox: {len(snap)} rows, {len(to_recover)} to recover, {len(analysis['archiver'])} to archive, "
        f"{review_count} to review, {remaining_count} left for a later run"
    )
    for row in analysis["a_verifier"] + waiting:
        log("  to review:", row["from"], "|", row["objet"][:70], json.dumps(verdicts.get(row["id"], {})))
    report: dict[str, Any] = {
        "examined": len(snap),
        "deposited": 0,
        "recovered": set(),
        "archived": 0,
        "waiting": review_count,
        "remaining": remaining_count,
    }
    if DRY_RUN:
        _log_dry_run_inbox(to_recover, analysis["archiver"])
        return report

    batch = InboxBatch(client=client, verdicts=verdicts, subjects=subjects, script=script, report=report)
    done = _deposit_candidates(batch, to_recover)
    for row in waiting:
        verdict = verdicts.get(row["id"], {})
        verdict_name = verdict.get("verdict")
        note = "" if verdict_name else "sans avis du modèle, nouvel essai au prochain passage"
        review_once(
            f"INBOX:{row['id']}:{verdict_name or ''}",
            subjects[row["id"]],
            PLACE["outlook"],
            detail(row, verdict, note),
            verdict,
        )
    for row in analysis["a_verifier"]:
        review_once(f"INBOX:{row['id']}:", subjects[row["id"]], PLACE["outlook"], detail(row, {}))
    report["archived"] = move(
        client,
        done + analysis["archiver"],
        names["archive"],
        lambda row: event(
            "archivé",
            subjects[row["id"]],
            PLACE["outlook"],
            PLACE["archive"],
            f"{row['from'] or '(sans adresse)'} : {row['motif']}",
        ),
    )
    log(
        f"inbox: {report['deposited']} deposited, {len(report['recovered'])} senders recovered, "
        f"{report['archived']} archived"
    )
    return report


def state() -> dict[str, Any]:
    """Read the previous outlook-state.json, returning an empty dict if absent or invalid."""
    try:
        return json.loads((DATA / "outlook-state.json").read_text())
    except OSError, ValueError:
        return {}


def notify(text: str) -> None:
    """Invoke ``$OUTLOOK_SORTER_NOTIFY`` if configured."""
    cmd = os.environ.get("OUTLOOK_SORTER_NOTIFY")
    if cmd:
        subprocess.run([cmd, "Tri mails Outlook", text], capture_output=True, timeout=60)


def save_waiting(previous: Any = ()) -> None:
    """Persist waiting mail keys atomically (mode 600) to suppress duplicate review events."""
    if DRY_RUN:
        return
    tmp = WAITING.with_name(f".{WAITING.name}.tmp")
    try:
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as handle:
            json.dump(sorted(waiting_now | set(previous)), handle)
        os.replace(tmp, WAITING)
    except OSError as error:
        log("waiting mails not kept:", clean(error, 120))


def fail(reason: str, public: bool, error: str | None = None) -> None:
    """Record a pass failure in state, journal, stdout summary, and optional notification."""
    if not DRY_RUN:
        try:
            run_script("record.py", ["--failure", reason])
        except (OSError, subprocess.SubprocessError, RuntimeError) as exc:
            log("failure not recorded:", clean(exc, 120))
    event("échec", "passage Outlook", detail=error or reason)
    save_waiting(reported)
    print(f"SUMMARY: {reason}", flush=True)
    notify(reason if public else "Le passage Outlook a échoué, voir son journal")


def timeout(*_: Any) -> None:
    """SIGALRM handler enforcing the 30-minute pass watchdog."""
    fail("Le passage Outlook dépasse 30 minutes : arrêté", True)
    os._exit(1)


def _record_pass_outcome(junk: dict[str, Any], box: dict[str, Any], to_review: int) -> None:
    """Invoke record.py with pass totals and journal any newly learned senders or spam domains."""
    save_waiting()
    per_domain = Counter(row["from"].rpartition("@")[2] for row in junk["new_spam"])
    out = run_script(
        "record.py",
        [
            "--examined",
            junk["examined"] + box["examined"],
            "--restored",
            len(junk["restored"]),
            "--deleted",
            junk["deleted"],
            "--to-review",
            to_review,
            "--recovered",
            box["deposited"],
            "--archived",
            box["archived"],
            "--remaining",
            box["remaining"],
            "--spam",
            *[domain for domain, count in per_domain.items() if count >= 3],
            "--legitimate",
            *junk["restored"],
        ],
    )
    try:
        learned = dict(json.loads(out))
    except ValueError, TypeError:
        learned = {}
    for address in learned.get("legitimate", []):
        event(
            "ajouté aux légitimes",
            address,
            detail="déjà reçu dans Gmail : ses mails indésirables seront remis en boîte sans avis du modèle",
        )
    for domain in learned.get("spam", []):
        event(
            "ajouté aux domaines spam",
            domain,
            detail=(
                f"{per_domain[domain]} spams sûrs de ce domaine à ce passage : "
                "ses prochains mails seront supprimés sans avis du modèle"
            ),
        )


def main() -> int:
    """Execute one full IMAP sorting pass and return 0 on success or 1 on failure."""
    EML.mkdir(parents=True, exist_ok=True, mode=0o700)
    signal.signal(signal.SIGALRM, timeout)
    signal.alarm(MAX_DURATION)
    before = state().get("a_verifier", 0)
    settings = json.loads((DATA / "senders.json").read_text())
    try:
        reported.update(json.loads(WAITING.read_text()))
    except OSError, ValueError, TypeError:
        pass
    client = None
    try:
        client = outlook_imap.connect()
        names = outlook_imap.folders(client)
        junk = junk_folder(client, names)
        box = inbox(client, names, settings.get("reinjecter"))
        to_review = junk["waiting"] + box["waiting"]
        if not DRY_RUN:
            _record_pass_outcome(junk, box, to_review)
        prefix = "essai, rien de changé. " if DRY_RUN else ""
        summary = (
            f"{prefix}{len(junk['restored'])} remis en boîte, {junk['deleted']} spams supprimés, "
            f"{box['deposited']} déposés dans Gmail, {box['archived']} archivés, {to_review} à vérifier."
        )
        print(f"SUMMARY: {summary}", flush=True)
        event("passage terminé", "Outlook.com", detail=summary)
        if not DRY_RUN and to_review > before:
            notify(summary)
        return 0
    except outlook_imap.AuthError as error:
        fail(str(error), True)
    except (
        OSError,
        imaplib.IMAP4.error,
        subprocess.SubprocessError,
        RuntimeError,
        ValueError,
        KeyError,
        TypeError,
    ) as error:
        log("failure:", type(error).__name__, clean(error, 200))
        fail(clean(error, 150), False, clean(f"{type(error).__name__} : {error}", 1000))
    finally:
        if client is not None:
            try:
                client.logout()
            except OSError, imaplib.IMAP4.error:
                pass
        for json_file in WORK.glob("*.json"):
            json_file.unlink(missing_ok=True)
    return 1


if __name__ == "__main__":
    sys.exit(main())
