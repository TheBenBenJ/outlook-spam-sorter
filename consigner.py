#!/opt/homebrew/bin/python3 -I
"""Consigne un passage de tri-spam-outlook dans expediteurs.json, approuves.json et etat-outlook.json.

Seules entrent dans `legitimes` les adresses déjà reçues dans Gmail, et jamais une messagerie grand public dans les domaines de spam : une ligne de spam n'y bannit pas tous ses utilisateurs.
consigner.py --examines N --remis N --supprimes N --a-verifier N [--recuperes N --archives N --restants N] [--spam domaine ...] [--legitimes adresse ...] [--approuves adresse ...]
consigner.py --echec "motif"
"""
import argparse, datetime, json, os, re
from pathlib import Path

ICI = Path(__file__).resolve().parent
MEMOIRE, ETAT = ICI / "expediteurs.json", ICI / "etat-outlook.json"

p = argparse.ArgumentParser()
for nom in ("--examines", "--remis", "--supprimes", "--a-verifier", "--recuperes", "--archives", "--restants"):
    p.add_argument(nom, type=int, default=0)
p.add_argument("--spam", nargs="*", default=[])
p.add_argument("--legitimes", nargs="*", default=[])
p.add_argument("--approuves", nargs="*", default=[])
p.add_argument("--echec", default=None)
a = p.parse_args()


def ecrire(ch, d):
    tmp = ch.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=2) + "\n")
    os.replace(tmp, ch)


maintenant = datetime.datetime.now().isoformat(timespec="minutes")
if a.echec:
    ancien = json.loads(ETAT.read_text()) if ETAT.exists() else {}
    ecrire(ETAT, {**ancien, "maj": maintenant, "echec": a.echec[:160]})
    print("échec consigné")
    raise SystemExit(0)

d = json.loads(MEMOIRE.read_text())
try:
    connus = json.loads(Path(d["gmail_connus"]).read_text())
except (KeyError, OSError, ValueError):
    connus = {"adresses": {}, "domaines": {}}
CH_APPROUVES = ICI / "approuves.json"
approuves = set(json.loads(CH_APPROUVES.read_text())) if CH_APPROUVES.exists() else set()
fournisseurs = set(connus.get("fournisseurs") or [])
for adresse in {x.lower() for x in a.legitimes}:
    domaine = adresse.rpartition("@")[2]
    if adresse in connus["adresses"] or (fournisseurs and domaine not in fournisseurs and connus["domaines"].get(domaine, 0) >= 5):
        d["legitimes"].setdefault(adresse, f"{datetime.date.today().isoformat()} déjà reçu dans Gmail")
        approuves.add(adresse)
approuves |= {x.lower() for x in a.approuves}
proteges = {k.lower().rpartition("@")[2] for k in d["legitimes"]}
ajouts = {s.lower() for s in a.spam if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", s.lower())} - proteges - fournisseurs
d["spam_confirme_domaines"] = sorted(set(d["spam_confirme_domaines"]) | ajouts)
d["historique"].append({"date": datetime.date.today().isoformat(), "mode": "auto", "examines": a.examines, "remis_en_boite": a.remis, "desabonnements": 0, "supprimes": a.supprimes,
                        "a_verifier": a.a_verifier, "recuperes": a.recuperes, "archives": a.archives})
d["historique"] = d["historique"][-120:]
ecrire(MEMOIRE, d)
ecrire(CH_APPROUVES, sorted(approuves))
ecrire(ETAT, {"maj": maintenant, "a_verifier": a.a_verifier, "recuperes": a.recuperes, "archives": a.archives, "restants": a.restants, "echec": None})
print("consigné")
