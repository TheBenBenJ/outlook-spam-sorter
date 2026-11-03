#!/usr/bin/env python3
"""Offline checks of imap-pass.py and outlook_imap.py: a fake IMAP server, a fake `claude` for classify.py and a fake deposit script.
Nothing reaches Outlook, Gmail or Claude. Runs on a temporary copy of the repository.   python3 -I tests.py
"""
import datetime, importlib.util, json, os, shutil, subprocess, sys, tempfile
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
    for f in ("analyze.py", "classify.py", "record.py", "imap-pass.py", "outlook_imap.py"):
        shutil.copy(REPO / f, tmp / f)
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

    ok(oi.mutf7("Courrier ind&AOk-sirable") == "Courrier indésirable" and oi.mutf7("A&-B") == "A&B", "folder names decoded")
    ok(oi.folders(fake) == {"junk": "Junk", "trash": "Deleted", "archive": "Archivage"}, "folders found by flag, the archive by name")

r = subprocess.run([sys.executable, "-I", str(REPO / "imap-pass.py"), "essai"], capture_output=True, text=True)
ok(r.returncode != 0 and "usage" in r.stderr, "the old dry-run word is refused")
print("all checks passed" if not failures else f"{failures} check(s) failed")
sys.exit(1 if failures else 0)
