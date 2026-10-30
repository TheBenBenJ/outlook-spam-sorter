#!/opt/homebrew/bin/python3 -I
"""Consigne un passage auto de tri-spam-outlook dans expediteurs.json : seule écriture permise à ce passage.

Ne touche jamais à `legitimes` : un mail piégé ne peut pas y faire entrer un expéditeur.
consigner.py --examines N --remis N --supprimes N --a-verifier N [--spam domaine ...]
"""
import argparse, datetime, json, os, re
from pathlib import Path

MEMOIRE = Path(__file__).resolve().parent / "expediteurs.json"

p = argparse.ArgumentParser()
for nom in ("--examines", "--remis", "--supprimes", "--a-verifier"):
    p.add_argument(nom, type=int, required=True)
p.add_argument("--spam", nargs="*", default=[])
a = p.parse_args()

d = json.loads(MEMOIRE.read_text())
proteges = {k.lower().rpartition("@")[2] for k in d["legitimes"]}
ajouts = {s.lower() for s in a.spam if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", s.lower())} - proteges
d["spam_confirme_domaines"] = sorted(set(d["spam_confirme_domaines"]) | ajouts)
d["historique"].append({"date": datetime.date.today().isoformat(), "mode": "auto", "examines": a.examines, "remis_en_boite": a.remis,
                        "desabonnements": 0, "supprimes": a.supprimes, "a_verifier": a.a_verifier})
tmp = MEMOIRE.with_suffix(".tmp")
tmp.write_text(json.dumps(d, ensure_ascii=False, indent=2) + "\n")
os.replace(tmp, MEMOIRE)
print("consigné")
