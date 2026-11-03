#!/opt/homebrew/bin/python3 -I
"""Record an outlook-spam-sorter run in senders.json, approved.json and outlook-state.json.

Only addresses already received in Gmail enter `legitimes`, and a consumer mail provider never enters the spam domains: one spam row
must not ban all of its users.
record.py --examined N --restored N --deleted N --to-review N [--recovered N --archived N --remaining N] [--spam domain ...] [--legitimate address ...] [--approved address ...]
record.py --failure "reason"
Prints what it learned, for the activity journal: {"legitimate": [addresses added to legitimes], "spam": [domains added to the spam domains]}.
Other components read these three files (mail sorting, the Dashboard, the relay to the home server): their keys and values stay as they are.
They are in $OUTLOOK_SORTER_DATA, by default this folder.
"""
import argparse, datetime, json, os, re
from pathlib import Path

DATA = Path(os.environ.get("OUTLOOK_SORTER_DATA") or Path(__file__).resolve().parent)
MEMORY, STATE = DATA / "senders.json", DATA / "outlook-state.json"

p = argparse.ArgumentParser()
for name in ("--examined", "--restored", "--deleted", "--to-review", "--recovered", "--archived", "--remaining"):
    p.add_argument(name, type=int, default=0)
p.add_argument("--spam", nargs="*", default=[])
p.add_argument("--legitimate", nargs="*", default=[])
p.add_argument("--approved", nargs="*", default=[])
p.add_argument("--failure", default=None)
a = p.parse_args()


def write(path, data):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    os.replace(tmp, path)


now = datetime.datetime.now().isoformat(timespec="minutes")
if a.failure:
    previous = json.loads(STATE.read_text()) if STATE.exists() else {}
    write(STATE, {**previous, "maj": now, "echec": a.failure[:160]})
    print("failure recorded")
    raise SystemExit(0)

mem = json.loads(MEMORY.read_text())
try:
    known = json.loads(Path(mem["gmail_connus"]).read_text())
except (KeyError, OSError, ValueError):
    known = {"adresses": {}, "domaines": {}}
APPROVED = DATA / "approved.json"
approved = set(json.loads(APPROVED.read_text())) if APPROVED.exists() else set()
providers = set(known.get("fournisseurs") or [])
learned = {"legitimate": [], "spam": []}
for address in sorted({x.lower() for x in a.legitimate}):
    domain = address.rpartition("@")[2]
    if address in known["adresses"] or (providers and domain not in providers and known["domaines"].get(domain, 0) >= 5):
        if address not in mem["legitimes"]:
            learned["legitimate"].append(address)
        mem["legitimes"].setdefault(address, f"{datetime.date.today().isoformat()} déjà reçu dans Gmail")  # a note people read, like the rest of the file: French
        approved.add(address)
approved |= {x.lower() for x in a.approved}
protected = {k.lower().rpartition("@")[2] for k in mem["legitimes"]}
additions = {s.lower() for s in a.spam if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", s.lower())} - protected - providers
learned["spam"] = sorted(additions - set(mem["spam_confirme_domaines"]))
mem["spam_confirme_domaines"] = sorted(set(mem["spam_confirme_domaines"]) | additions)
mem["historique"].append({"date": datetime.date.today().isoformat(), "mode": "auto", "examines": a.examined, "remis_en_boite": a.restored, "desabonnements": 0, "supprimes": a.deleted,
                          "a_verifier": a.to_review, "recuperes": a.recovered, "archives": a.archived})
mem["historique"] = mem["historique"][-120:]
write(MEMORY, mem)
write(APPROVED, sorted(approved))
write(STATE, {"maj": now, "a_verifier": a.to_review, "recuperes": a.recovered, "archives": a.archived, "restants": a.remaining, "echec": None})
print(json.dumps(learned))
