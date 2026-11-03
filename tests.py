#!/usr/bin/env python3
"""Offline checks of imap-pass.py, outlook_imap.py and activity.py: a fake IMAP server, a fake `claude` for classify.py and a fake deposit
script. Nothing reaches Outlook, Gmail or Claude. Runs on a temporary copy of the repository.   python3 -I tests.py
"""
import datetime, importlib.util, json, os, shutil, subprocess, sys, tempfile
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent
TODAY = datetime.date.today()
failures = 0


def ok(cond, what):
    global failures
    print(("ok   " if cond else "FAIL ") + what)
    failures += not cond


class FakeIMAP:
    """Outlook.com as seen over IMAP: Junk and Deleted carry special-use flags, the archive folder does not."""
    def __init__(self, mails):
        self.boxes = {"INBOX": {}, "Junk": {}, "Deleted": {}, "Archivage": {}, "Sent": {}}
        self.next_uid, self.selected, self.readonly, self.log = 1, None, None, []
        for box, sender, subject, day in mails:
            self.add(box, sender, subject, day)

    def add(self, box, sender, subject, day):
        self.boxes[box][str(self.next_uid)] = (sender, subject, day)
        self.next_uid += 1

    def list(self):
        return "OK", [b'(\\HasNoChildren) "/" Archivage', b'(\\Marked \\HasNoChildren) "/" Inbox', b'(\\HasNoChildren \\Junk) "/" Junk',
                      b'(\\HasNoChildren \\Trash) "/" Deleted', b'(\\HasNoChildren \\Sent) "/" Sent']

    def select(self, name, readonly=False):
        self.selected, self.readonly = name.strip('"'), readonly
        return "OK", [str(len(self.boxes[self.selected])).encode()]

    def uid(self, cmd, *a):
        self.log.append((cmd, self.selected, self.readonly, a))
        box = self.boxes[self.selected]
        if cmd == "SEARCH":
            return "OK", [" ".join(box).encode()]
        if cmd == "FETCH":
            out = []
            for u in a[0].split(","):
                sender, subject, day = box[u]
                d = datetime.date.fromisoformat(day).strftime("%d-%b-%Y")
                head = f"From: {sender}\r\nSubject: {subject}\r\n\r\n".encode()
                body = head + b"Hello" if "BODY.PEEK[]" in a[1] else head
                out += [(f'{u} (UID {u} INTERNALDATE "{d} 10:00:00 +0000" BODY[] {{{len(body)}}}'.encode(), body), b")"]
            return "OK", out
        if cmd == "MOVE":
            assert not self.readonly, "MOVE in a read-only folder"
            for u in a[0].split(","):
                self.add(a[1].strip('"'), *box.pop(u))
            return "OK", [b"done"]
        raise AssertionError(f"unexpected IMAP command {cmd}")

    def store(self, *a):
        raise AssertionError("STORE (a deletion flag) must never be used")

    def expunge(self):
        raise AssertionError("EXPUNGE must never be used")

    def logout(self):
        pass

    def senders(self, box):
        return sorted(x[0].split("<")[-1].rstrip(">") for x in self.boxes[box].values())


def setup(tmp):
    for f in ("activity.py", "analyze.py", "classify.py", "record.py", "imap-pass.py", "outlook_imap.py"):
        shutil.copy(REPO / f, tmp / f)
    os.environ.pop("OUTLOOK_SORTER_ACTIVITY", None)   # the journal goes next to the state, in the temporary copy
    old = (TODAY - datetime.timedelta(days=200)).isoformat()
    (tmp / "known.json").write_text(json.dumps({"adresses": {"friend@example.org": 4}, "domaines": {"example.org": 4}, "fournisseurs": ["gmail.com"],
                                                "libelles": {}, "exacts": {}, "recents": []}))
    (tmp / "senders.json").write_text(json.dumps({"legitimes": {"bank@example.net": "bank"}, "spam_confirme_domaines": ["spam.example"], "desabonnes": {},
                                                  "historique": [], "moi": ["me@example.com"], "gmail_connus": str(tmp / "known.json"), "libelles_bruit": [],
                                                  "recuperer_depuis": "2000-01-01", "reinjecter": [sys.executable, "-I", str(tmp / "deposit.py")]}))
    (tmp / "deposit.py").write_text("import json, sys\nopen(sys.argv[0] + '.log', 'a').write(open(sys.argv[1], 'rb').read().decode() + '\\n---\\n')\n"
                                    "print(json.dumps({'statut': 'depose', 'de': 'x', 'objet': 'y'}))\n")
    (tmp / "bin").mkdir()
    # fake claude: "spam" for random-looking domains, "legitime" otherwise, read from the rows in the prompt
    (tmp / "bin" / "claude").write_text("#!/usr/bin/env python3\nimport json, re, sys\nrows = [json.loads(l) for l in sys.argv[2].splitlines() if l.startswith('{')]\n"
                                        "print(json.dumps({'structured_output': {'avis': [{'n': r['n'], 'verdict': 'spam' if 'xq' in r['from'] else 'legitime', "
                                        "'confiance': 0.95} for r in rows]}}))\n")
    os.chmod(tmp / "bin" / "claude", 0o755)
    os.environ["PATH"] = f"{tmp / 'bin'}:{os.environ['PATH']}"
    os.environ["OUTLOOK_SORTER_LOCAL_URL"] = "off"   # Haiku (the fake claude) first; the local model has its own checks below
    today = TODAY.isoformat()
    return FakeIMAP([
        ("Junk", "Bank <bank@example.net>", "Votre relevé", today),               # listed as legitimate: back to the inbox, then deposited
        ("Junk", "friend@example.org", "Photos du week-end", today),            # known in Gmail, judged legitimate: same
        ("Junk", "Casino <a@spam.example>", "BONUS 400%", today),               # known spam domain: deleted
        ("Junk", "Colis <no-reply@xq7rt.biz>", "Votre colis", today),           # unknown, judged spam: deleted
        ("Junk", "Shop <news@shop.example>", "Nouveautés", today),              # unknown, judged legitimate: waits for a person
        ("INBOX", "me@example.com", "Note à moi", today),                       # my own address: archived
        ("INBOX", "stranger@unknown.example", "Bonjour", today),                # unknown and recent: waits
        ("INBOX", "old@unknown.example", "Ancien", old),                        # unknown and old: archived
    ])


def load(tmp, argv):
    sys.argv = ["imap-pass.py", *argv]
    sys.path.insert(0, str(tmp))
    spec = importlib.util.spec_from_file_location(f"imap_pass_{len(argv)}_{id(tmp)}", tmp / "imap-pass.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run(tmp, fake, argv=(), connect=None, notes=None):
    mod = load(tmp, list(argv))
    mod.outlook_imap.connect = connect or (lambda: fake)
    mod.notify = (lambda text: notes.append(text)) if notes is not None else (lambda text: None)
    return mod.main()


with tempfile.TemporaryDirectory() as d:
    tmp = Path(d)
    fake = setup(tmp)
    ok(run(tmp, fake, ["trial"]) == 0, "trial pass ends well")
    ok(all(not c[0] == "MOVE" for c in fake.log) and all(c[2] for c in fake.log), "trial: folders read-only, nothing moved")
    ok(not (tmp / "outlook-state.json").exists(), "trial: nothing recorded")
    journal = tmp / "activity.jsonl"
    ok(not journal.exists() and not (tmp / "activity-waiting.json").exists(), "trial: no activity journal")

    notes = []
    ok(run(tmp, fake, notes=notes) == 0, "real pass ends well")
    ok(fake.senders("Deleted") == ["a@spam.example", "no-reply@xq7rt.biz"], "spam to Deleted Items: known domain and confident verdict")
    ok(fake.senders("Junk") == ["news@shop.example"], "an unknown sender judged legitimate stays in junk")
    ok(fake.senders("Archivage") == ["bank@example.net", "friend@example.org", "me@example.com", "old@unknown.example"], "deposited, own and old mails archived")
    ok(fake.senders("INBOX") == ["stranger@unknown.example"], "a recent unknown sender stays in the inbox")
    deposited = (tmp / "deposit.py.log").read_text()
    ok("bank@example.net" in deposited and "friend@example.org" in deposited and deposited.count("---") == 2, "the two trusted mails deposited, original .eml")
    st = json.loads((tmp / "outlook-state.json").read_text())
    ok(st["a_verifier"] == 2 and st["recuperes"] == 2 and st["archives"] == 4 and st["echec"] is None, "outlook-state.json written")
    mem = json.loads((tmp / "senders.json").read_text())
    ok("friend@example.org" in mem["legitimes"] and mem["historique"][-1]["supprimes"] == 2, "senders.json: known sender learned, run in the history")
    ok(len(notes) == 1 and "2 à vérifier" in notes[0] and "@" not in notes[0], "notified once (more to review), without any address")
    ok(not list((tmp / "work" / "eml").iterdir()) and not list((tmp / "work").glob("*.json")), "work/ emptied")

    events = [json.loads(x) for x in journal.read_text().splitlines()]
    ok(Counter(e["action"] for e in events) == {"remis en boîte": 2, "spam supprimé": 2, "à vérifier": 2, "récupéré dans Gmail": 2, "archivé": 4,
                                                "ajouté aux légitimes": 1, "passage terminé": 1}, "journal: one event per mail acted on, and the end of the pass")
    ok([e["object"] for e in events if e["action"] == "ajouté aux légitimes"] == ["friend@example.org"], "journal: a sender learned by record.py")
    ev = {(e["action"], e.get("detail", "").split(" : ")[0]): e for e in events}
    ok(ev["remis en boîte", "bank@example.net"] == {**ev["remis en boîte", "bank@example.net"], "object": "Votre relevé", "from": "Courrier indésirable",
                                                    "to": "Boîte de réception", "detail": "bank@example.net : adresse reconnue comme légitime"}
       and "engine" not in ev["remis en boîte", "bank@example.net"], "journal: a listed sender restored, without a model")
    e = ev["remis en boîte", "friend@example.org"]
    ok(e["engine"] == "claude" and e["model"] == "haiku" and e["confidence"] == 0.95 and e["detail"].endswith("expéditeur déjà reçu dans Gmail, jugé légitime (sûr à 95 %)"),
       "journal: a known sender restored on the model's verdict, with engine, model and confidence")
    ok(ev["spam supprimé", "a@spam.example"]["detail"].endswith("domaine connu pour le spam") and ev["spam supprimé", "a@spam.example"]["to"] == "Éléments supprimés"
       and ev["spam supprimé", "no-reply@xq7rt.biz"]["model"] == "haiku" and "jugé spam" in ev["spam supprimé", "no-reply@xq7rt.biz"]["detail"], "journal: spam deleted, and why")
    ok(ev["à vérifier", "news@shop.example"]["from"] == "Courrier indésirable" and "seul un expéditeur connu" in ev["à vérifier", "news@shop.example"]["detail"]
       and ev["à vérifier", "stranger@unknown.example"]["detail"].endswith("expéditeur inconnu"), "journal: mails left to review, and why")
    ok(ev["récupéré dans Gmail", "friend@example.org"]["from"] == "Boîte de réception Outlook" and ev["récupéré dans Gmail", "friend@example.org"]["to"] == "Gmail",
       "journal: stuck mail handed to Gmail")
    ok(ev["archivé", "me@example.com"]["detail"].endswith("envoyé depuis une de mes adresses") and ev["archivé", "old@unknown.example"]["detail"].endswith("de plus de 90 jours")
       and ev["archivé", "bank@example.net"]["detail"].endswith("récupéré dans Gmail") and ev["archivé", "bank@example.net"]["to"] == "Archivage", "journal: archived, and why")
    ok(ev["passage terminé", "2 remis en boîte, 2 spams supprimés, 2 déposés dans Gmail, 4 archivés, 2 à vérifier."]["object"] == "Outlook.com", "journal: the pass's counts")
    ok(len({e["run"] for e in events}) == 1 and all(datetime.datetime.fromisoformat(e["time"]).utcoffset() is not None and "" not in e.values() for e in events)
       and journal.stat().st_mode & 0o777 == 0o600, "journal: one run, times with their offset, no empty field, mode 600")
    n = len(events)
    ok(run(tmp, fake) == 0 and [json.loads(x)["action"] for x in journal.read_text().splitlines()[n:]] == ["passage terminé"],
       "journal: a mail still waiting is not reported again at the next pass")

    notes = []
    import imaplib  # noqa: E402
    from importlib import import_module  # noqa: E402
    oi = import_module("outlook_imap")

    def expired():
        raise oi.AuthError(oi.EXPIRED)
    ok(run(tmp, fake, connect=expired, notes=notes) == 1, "expired grant: the run fails")
    st = json.loads((tmp / "outlook-state.json").read_text())
    ok(st["echec"].startswith("Session Outlook") and notes == [oi.EXPIRED], "expired grant: Dashboard reason and notification")

    def broken():
        raise RuntimeError("déplacement refusé par Outlook : friend@example.org")
    notes = []
    ok(run(tmp, fake, connect=broken, notes=notes) == 1 and notes == ["Le passage Outlook a échoué, voir son journal"], "other failure: generic notification, no address")
    failures_logged = [json.loads(x) for x in journal.read_text().splitlines()][-2:]
    ok([(e["action"], e["object"]) for e in failures_logged] == [("échec", "passage Outlook")] * 2 and failures_logged[0]["detail"].startswith("Session Outlook")
       and failures_logged[1]["detail"] == "RuntimeError : déplacement refusé par Outlook : friend@example.org", "journal: each failed pass, with its error")

    # the deposit refuses one mail and fails for another: the first waits for a person, the second is retried at the next pass
    (tmp / "deposit.py").write_text("import json, sys\nraw = open(sys.argv[1], 'rb').read()\nprint(json.dumps({'statut': 'refuse', 'motif': 'DMARC en échec'} "
                                    "if b'bank@' in raw else {'statut': 'echec', 'motif': 'Gmail injoignable'}))\n")
    fake.add("INBOX", "Bank <bank@example.net>", "Relevé suspect", TODAY.isoformat())
    fake.add("INBOX", "friend@example.org", "Encore des photos", TODAY.isoformat())
    for i in range(3):   # three confident spams from one new domain: record.py learns it
        fake.add("Junk", f"promo{i}@xqz.example", "Gagnez !", TODAY.isoformat())
    n = len(journal.read_text().splitlines())
    ok(run(tmp, fake) == 0, "pass with a refused and a failed deposit ends well")
    new = [json.loads(x) for x in journal.read_text().splitlines()[n:]]
    ok([e["action"] for e in new] == ["spam supprimé"] * 3 + ["à vérifier", "échec", "ajouté aux domaines spam", "passage terminé"]
       and new[3]["object"] == "Relevé suspect" and new[3]["detail"].endswith("dépôt dans Gmail refusé : DMARC en échec") and new[4]["object"] == "Encore des photos"
       and new[4]["to"] == "Gmail" and "dépôt dans Gmail échoué : Gmail injoignable" in new[4]["detail"],
       "journal: a refused deposit waits, a failed one is an « échec » on that mail")
    ok(new[5]["object"] == "xqz.example" and new[5]["detail"].startswith("3 spams sûrs"), "journal: a spam domain learned by record.py")
    shadow = load(tmp, []).detail({"from": "a@b.example", "motif": "expéditeur inconnu"},
                                  {"verdict": "douteux", "confiance": 0.5, "engine": "claude", "model": "haiku", "local": {"verdict": "legitime", "confiance": 0.9}})
    ok(shadow == "a@b.example : expéditeur inconnu, jugé douteux (sûr à 50 %) ; le modèle local disait légitime (sûr à 90 %)",
       "journal, shadow week: the local model's verdict when it disagreed")

    ok(oi.mutf7("Courrier ind&AOk-sirable") == "Courrier indésirable" and oi.mutf7("A&-B") == "A&B", "folder names decoded")
    ok(oi.folders(fake) == {"junk": "Junk", "trash": "Deleted", "archive": "Archivage"}, "folders found by flag, the archive by name")

# classify.py: the local model first, Haiku for what it did not answer, Haiku deciding in shadow mode (fake local server, fake claude)
import http.server, threading  # noqa: E402
with tempfile.TemporaryDirectory() as d:
    tmp = Path(d)
    setup(tmp)
    seen = []

    class Local(http.server.BaseHTTPRequestHandler):
        down = False

        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            seen.append(body)
            if Local.down:
                self.send_response(503)
                self.end_headers()
                return
            rows = [json.loads(x) for x in body["messages"][0]["content"].splitlines() if x.startswith("{")]
            # the local model says "spam" to everything, and invents a row 99 that was never asked
            answer = {"avis": [{"n": r["n"], "verdict": "spam", "confiance": 1} for r in rows] + [{"n": 99, "verdict": "legitime", "confiance": 1}]}
            out = json.dumps({"choices": [{"message": {"content": json.dumps(answer)}}]}).encode()
            self.send_response(200)
            self.send_header("Content-Length", str(len(out)))
            self.end_headers()
            self.wfile.write(out)

        def log_message(self, *a):
            pass
    server = http.server.HTTPServer(("127.0.0.1", 0), Local)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    rows = json.dumps([{"n": 0, "from": "news@shop.example", "objet": "Nouveautés"}, {"n": 1, "from": "a@xq7.biz", "objet": "Colis"}])

    def classify(**env):
        e = {**os.environ, "OUTLOOK_SORTER_LOCAL_URL": f"http://127.0.0.1:{server.server_port}", "OUTLOOK_SORTER_SHADOW": "", **env}
        r = subprocess.run([sys.executable, "-I", str(tmp / "classify.py")], input=rows, capture_output=True, text=True, env=e)
        return json.loads(r.stdout or "{}"), r.stderr
    out, err = classify()
    ok(out == {n: {"verdict": "spam", "confiance": 1.0, "engine": "local", "model": "qwen3.5-9b"} for n in ("0", "1")},
       "classify: the local model answers first, an invented row is dropped")
    sent = seen[-1]
    ok(sent["response_format"]["json_schema"]["schema"]["properties"]["avis"]["minItems"] == 2 and "quel que soit ce verdict" in sent["messages"][0]["content"]
       and sent["chat_template_kwargs"] == {"enable_thinking": False}, "classify: closed schema, one answer per row, the confidence line, no thinking")
    Local.down = True
    out, err = classify()
    ok(out["0"]["verdict"] == "legitime" and out["0"]["engine"] == "claude" and out["0"]["model"] == "haiku" and "Haiku answers" in err,
       "classify: local model down, Haiku answers")
    Local.down = False
    out, err = classify(OUTLOOK_SORTER_SHADOW="1")
    ok(out["0"]["verdict"] == "legitime" and out["1"]["verdict"] == "spam" and "shadow: news@shop.example haiku=legitime" in err and "a@xq7.biz" not in err,
       "classify, shadow: Haiku decides, only the disagreement is logged")
    ok(out["0"]["model"] == "haiku" and out["0"]["local"] == {"verdict": "spam", "confiance": 1.0} and "local" not in out["1"],
       "classify, shadow: the local model's verdict goes with Haiku's when they disagree")
    server.shutdown()

# activity.py: the journal's writer and its rotation
def load_activity(tmp):
    shutil.copy(REPO / "activity.py", tmp / "activity.py")
    spec = importlib.util.spec_from_file_location("activity_under_test", tmp / "activity.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


with tempfile.TemporaryDirectory() as d:
    tmp = Path(d)
    os.environ["OUTLOOK_SORTER_ACTIVITY"] = str(tmp / "journal.jsonl")
    act = load_activity(tmp)
    ok(act.PATH == tmp / "journal.jsonl", "activity: $OUTLOOK_SORTER_ACTIVITY sets the path")
    del os.environ["OUTLOOK_SORTER_ACTIVITY"]
    ok(load_activity(tmp).PATH == tmp.resolve() / "activity.jsonl", "activity: by default next to the pass's state")
    act.log("spam supprimé", "x" * 400, **{"from": "Courrier indésirable", "to": ""}, detail="d" * 1200, engine=None, confidence=0.0, run="r")
    act.log("archivé", "a\u2028b")
    lines = act.PATH.read_text(encoding="utf-8").splitlines()
    e = json.loads(lines[0])
    ok(set(e) == {"time", "action", "object", "from", "detail", "confidence", "run"} and len(e["object"]) == 300 and len(e["detail"]) == 1000
       and datetime.datetime.fromisoformat(e["time"]).utcoffset() is not None and act.PATH.stat().st_mode & 0o777 == 0o600,
       "activity: empty fields left out, object and detail cut, time with its offset, mode 600")
    ok(len(lines) == 2 and json.loads(lines[1])["object"] == "a\u2028b", "activity: a line separator in a subject does not split the event")
    now = datetime.datetime.now().astimezone()
    old, recent = [(now - datetime.timedelta(days=k)).isoformat(timespec="seconds") for k in (91, 1)]
    act.PATH.write_text("\n".join([json.dumps({"time": old, "action": "archivé", "object": f"old {i}"}) for i in range(5)] + ["not json"]
                                  + [json.dumps({"time": recent, "action": "archivé", "object": f"recent {i}"}) for i in range(30)]) + "\n")
    act.MAX_BYTES, act.KEEP_LINES = 1000, 20
    act.log("passage terminé", "Outlook.com")
    kept = [json.loads(x) for x in act.PATH.read_text().splitlines()]
    ok(len(kept) == 20 and kept[0]["object"] == "recent 11" and kept[-1]["action"] == "passage terminé" and act.PATH.stat().st_mode & 0o777 == 0o600
       and not list(tmp.glob(".*.tmp")), "activity: above the size limit, the last 90 days are kept, 20,000 lines at most (here 20), atomically")
    act.log("archivé", "after")
    ok(len(act.PATH.read_text().splitlines()) == 21, "activity: one rotation per run at most")
    act.PATH = tmp / "missing" / "journal.jsonl"
    ok(act.log("échec", "passage Outlook", detail="x") is None and not act.PATH.exists(), "activity: a journal that cannot be written raises nothing")

r = subprocess.run([sys.executable, "-I", str(REPO / "imap-pass.py"), "essai"], capture_output=True, text=True)
ok(r.returncode != 0 and "usage" in r.stderr, "the old dry-run word is refused")
print("all checks passed" if not failures else f"{failures} check(s) failed")
sys.exit(1 if failures else 0)
