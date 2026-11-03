#!/usr/bin/env python3
"""Daily pass over Outlook.com by IMAP (XOAUTH2): the same pass as daily-pass.mjs, without a browser. Mail content is hostile: what this
script does never depends on what a mail says. It opens no link, sends nothing, never deletes for good (deleting is a move to Deleted Items)
and never expunges.

imap-pass.py [trial]   trial = dry run: folders opened read-only, nothing moved, deposited or recorded

Junk folder: known senders go back to the inbox, sure spam goes to Deleted Items, the rest waits for a person (analyze.py, classify.py).
Inbox: the original of each stuck mail from a trusted sender is deposited into Gmail (`reinjecter` of senders.json), then archived;
what Gmail already has, noise and old unknown mail are archived. Then record.py, as the browser pass does.
Not done over IMAP: Outlook's safe senders list (no IMAP or Graph API for it): recovered senders are no longer added to it.
Optional: $OUTLOOK_SORTER_NOTIFY, a command called as `<cmd> "Tri mails Outlook" "<text>"` when the run fails or the number of mails to
review goes up. The text never carries an address or a subject: the server's notifications go through a public service.
"""
import datetime, email, email.header, email.utils, imaplib, json, os, re, signal, subprocess, sys, time
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import outlook_imap  # noqa: E402

WORK, EML = HERE / "work", HERE / "work" / "eml"
SPAM_CONFIDENCE, MAX_DURATION, CHUNK = 0.8, 30 * 60, 200
args = sys.argv[1:]
if args not in ([], ["trial"]):
    sys.exit("usage: imap-pass.py [trial]")
DRY_RUN = args == ["trial"]


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


def clean(s, n=300):
    s = re.sub(r"[\x00-\x1f\x7f​-‏‪-‮⁦-⁩]", " ", str(s or ""))
    return re.sub(r"\s+", " ", s).strip()[:n]


def run_script(name, script_args=(), stdin=None):
    r = subprocess.run([sys.executable, "-I", str(HERE / name), *map(str, script_args)], input=stdin, capture_output=True, text=True, timeout=900)
    if r.returncode != 0:
        raise RuntimeError(f"{name} : {(r.stderr or r.stdout).strip()[-200:]}")
    return r.stdout


def analyze(file, snapshot, analyze_args):
    (WORK / file).write_text(json.dumps(snapshot, ensure_ascii=False))
    return json.loads(run_script("analyze.py", [*analyze_args, file]))


def judge(rows):
    """The model's verdict for each row (classify.py: local model, else Haiku), keyed by id. Without a verdict nothing moves: the row waits for the next run."""
    if not rows:
        return {}
    try:
        r = subprocess.run([sys.executable, "-I", str(HERE / "classify.py")], input=json.dumps([{"n": n, "from": x["from"], "objet": x["objet"]} for n, x in enumerate(rows)]),
                           capture_output=True, text=True, timeout=900)
        for line in r.stderr.splitlines():   # which model answered, and in shadow mode each disagreement: kept in the journal
            log(" ", clean(line, 240))
        if r.returncode != 0:
            raise RuntimeError(f"classify.py : {r.stderr.strip()[-200:]}")
        out = json.loads(r.stdout)
        return {x["id"]: out.get(str(n), {}) for n, x in enumerate(rows)}
    except Exception as e:
        log("Haiku verdicts unavailable:", clean(e, 200))
        return {}


def header(msg, name):
    try:
        return str(email.header.make_header(email.header.decode_header(msg.get(name, ""))))
    except Exception:
        return msg.get(name, "")


def check(typ, data, what):
    if typ != "OK":
        raise RuntimeError(f"{what} refusé par Outlook : {clean(data, 120)}")
    return data


def snapshot(m, folder):
    """The folder's mails as analyze.py reads them: [{"id": UID, "from", "l": name and subject, "date": received date}]."""
    check(*m.select(f'"{folder}"', readonly=DRY_RUN), f"ouverture de {folder}")
    uids = check(*m.uid("SEARCH", None, "ALL"), "recherche")[0].split()
    rows = []
    for i in range(0, len(uids), CHUNK):
        data = check(*m.uid("FETCH", b",".join(uids[i:i + CHUNK]).decode(), "(UID INTERNALDATE BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])"), "lecture des en-têtes")
        for part in data:
            if not isinstance(part, tuple):
                continue
            uid = re.search(rb"UID (\d+)", part[0])
            when = imaplib.Internaldate2tuple(part[0])
            msg = email.message_from_string(part[1].decode("utf-8", "replace"))   # spam often puts raw UTF-8 in headers
            name, address = email.utils.parseaddr(header(msg, "From"))
            if uid:
                rows.append({"id": uid[1].decode(), "from": address.strip().lower(), "l": clean(f"{name} {header(msg, 'Subject')}"),
                             "date": time.strftime("%Y-%m-%d", when) if when else None})
    return rows


def move(m, uids, dest):
    """Moves by UID, in chunks. -> number moved. Never a deletion flag nor an expunge: Deleted Items keeps a mail 30 days."""
    done = 0
    for i in range(0, len(uids), CHUNK):
        chunk = uids[i:i + CHUNK]
        check(*m.uid("MOVE", ",".join(chunk), f'"{dest}"'), "déplacement")
        done += len(chunk)
    return done


def junk_folder(m, names):
    snap = snapshot(m, names["junk"])
    by_class = analyze("junk.json", snap, ["--lines"])["lignes"]
    to_judge = by_class["connu_gmail"] + by_class["a_classer"]
    verdicts = judge(to_judge)
    v = lambda x: verdicts.get(x["id"], {})  # noqa: E731
    restore = by_class["legitime"] + [x for x in by_class["connu_gmail"] if v(x).get("verdict") == "legitime"]
    new_spam = [x for x in by_class["a_classer"] if v(x).get("verdict") == "spam" and v(x).get("confiance", 0) >= SPAM_CONFIDENCE]
    remove = by_class["spam_connu"] + new_spam
    waiting = [x for x in to_judge if x not in restore and x not in remove]
    log(f"junk: {len(snap)} rows, {len(restore)} to restore, {len(remove)} spam, {len(waiting)} to review")
    for x in waiting:
        log("  to review:", x["from"], "|", x["objet"][:70], json.dumps(v(x)))
    report = {"examined": len(snap), "restored": [], "deleted": 0, "waiting": len(waiting), "new_spam": new_spam}
    if DRY_RUN:
        for x in restore:
            log("  to restore:", x["from"], "|", x["objet"][:70])
        for x in remove:
            log("  spam:", x["from"], "|", x["objet"][:70])
        return report
    move(m, [x["id"] for x in restore], "INBOX")   # back to the inbox, where the inbox sweep below recovers them into Gmail
    report["restored"] = [x["from"] for x in restore]
    report["deleted"] = move(m, [x["id"] for x in remove], names["trash"])
    log(f"junk: {len(restore)} restored, {report['deleted']} deleted")
    return report


def deposit(m, uid, script):
    """Original .eml of one mail -> the reinjection script. -> its JSON answer (statut: depose, deja, refuse or echec)."""
    data = check(*m.uid("FETCH", uid, "(BODY.PEEK[])"), "lecture du mail")
    raw = next((p[1] for p in data if isinstance(p, tuple)), None)
    if not raw:
        return {"statut": "echec", "motif": "mail introuvable"}
    f = EML / f"{uid}.eml"
    f.write_bytes(raw)
    cmd = script if isinstance(script, list) else [script]
    r = subprocess.run([*cmd, str(f)], capture_output=True, text=True, timeout=300)
    last = (r.stdout.strip().splitlines() or [""])[-1]
    try:
        return json.loads(last)
    except ValueError:
        return {"statut": "echec", "motif": clean(r.stderr, 120) or "pas de réponse"}
    finally:
        f.unlink(missing_ok=True)


def inbox(m, names, script):
    snap = snapshot(m, "INBOX")
    a = analyze("inbox.json", snap, ["inbox"])
    verdicts = judge(a["recuperer"])
    to_recover = [x for x in a["recuperer"] if verdicts.get(x["id"], {}).get("verdict") == "legitime"]
    waiting = [x for x in a["recuperer"] if x not in to_recover]
    log(f"inbox: {len(snap)} rows, {len(to_recover)} to recover, {len(a['archiver'])} to archive, {a['a_verifier_total'] + len(waiting)} to review, "
        f"{a['recuperer_restants'] + a['archiver_restants']} left for a later run")
    for x in a["a_verifier"] + waiting:
        log("  to review:", x["from"], "|", x["objet"][:70], json.dumps(verdicts.get(x["id"], {})))
    report = {"examined": len(snap), "deposited": 0, "recovered": set(), "archived": 0, "waiting": a["a_verifier_total"] + len(waiting),
              "remaining": a["recuperer_restants"] + a["archiver_restants"]}
    if DRY_RUN:
        for x in to_recover:
            log("  to recover:", x["from"], "|", x["objet"][:70])
        for x in a["archiver"]:
            log("  to archive:", x["from"], "|", x["objet"][:70])
        return report
    done = []
    for x in to_recover:
        r = deposit(m, x["id"], script)
        if r.get("statut") == "refuse":
            report["waiting"] += 1
            log("  refused:", r.get("de"), "|", clean(r.get("objet"), 50), r.get("motif"))
        elif r.get("statut") in ("depose", "deja"):
            report["recovered"].add(x["from"])
            report["deposited"] += r["statut"] == "depose"
            done.append(x["id"])
        else:
            report["remaining"] += 1
            log("  deposit failed:", r.get("de"), r.get("motif"))
    report["archived"] = move(m, done + [x["id"] for x in a["archiver"]], names["archive"])
    log(f"inbox: {report['deposited']} deposited, {len(report['recovered'])} senders recovered, {report['archived']} archived")
    return report


def state():
    try:
        return json.loads((HERE / "outlook-state.json").read_text())
    except (OSError, ValueError):
        return {}


def notify(text):
    cmd = os.environ.get("OUTLOOK_SORTER_NOTIFY")
    if cmd:
        subprocess.run([cmd, "Tri mails Outlook", text], capture_output=True, timeout=60)


def fail(reason, public):
    """The reason goes to outlook-state.json and the Dashboard (French). The notification only carries reasons known to hold nothing personal."""
    if not DRY_RUN:
        try:
            run_script("record.py", ["--failure", reason])
        except Exception as e:
            log("failure not recorded:", clean(e, 120))
    print(f"SUMMARY: {reason}", flush=True)
    notify(reason if public else "Le passage Outlook a échoué, voir son journal")


def timeout(*_):
    fail("Le passage Outlook dépasse 30 minutes : arrêté", True)
    os._exit(1)


def main():
    EML.mkdir(parents=True, exist_ok=True, mode=0o700)
    signal.signal(signal.SIGALRM, timeout)
    signal.alarm(MAX_DURATION)
    before = state().get("a_verifier", 0)
    settings = json.loads((HERE / "senders.json").read_text())
    m = None
    try:
        m = outlook_imap.connect()
        names = outlook_imap.folders(m)
        junk = junk_folder(m, names)
        box = inbox(m, names, settings.get("reinjecter"))
        to_review = junk["waiting"] + box["waiting"]
        if not DRY_RUN:
            per_domain = Counter(x["from"].rpartition("@")[2] for x in junk["new_spam"])
            run_script("record.py", ["--examined", junk["examined"] + box["examined"], "--restored", len(junk["restored"]), "--deleted", junk["deleted"],
                                     "--to-review", to_review, "--recovered", box["deposited"], "--archived", box["archived"], "--remaining", box["remaining"],
                                     "--spam", *[d for d, n in per_domain.items() if n >= 3], "--legitimate", *junk["restored"]])
        summary = (f"{'essai, rien de changé. ' if DRY_RUN else ''}{len(junk['restored'])} remis en boîte, {junk['deleted']} spams supprimés, "
                   f"{box['deposited']} déposés dans Gmail, {box['archived']} archivés, {to_review} à vérifier.")
        print(f"SUMMARY: {summary}", flush=True)
        if not DRY_RUN and to_review > before:
            notify(summary)
        return 0
    except outlook_imap.AuthError as e:
        fail(str(e), True)
    except Exception as e:
        log("failure:", type(e).__name__, clean(e, 200))
        fail(clean(e, 150), False)
    finally:
        if m is not None:
            try:
                m.logout()
            except Exception:
                pass
        for f in WORK.glob("*.json"):
            f.unlink(missing_ok=True)
    return 1


if __name__ == "__main__":
    sys.exit(main())
