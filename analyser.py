#!/opt/homebrew/bin/python3 -I
"""Regroupe par domaine le relevé du dossier indésirable : seule analyse permise au passage auto de tri-spam-outlook.

analyser.py <relevé.json>   liste de {"id", "from", "l"}, dans le dossier travail/ uniquement
"""
import json, sys
from collections import defaultdict
from pathlib import Path

ICI = Path(__file__).resolve().parent
TRAVAIL = (ICI / "travail").resolve()

if len(sys.argv) != 2 or (TRAVAIL / sys.argv[1]).resolve().parent != TRAVAIL:
    sys.exit("usage : analyser.py <fichier du dossier travail/>")
memoire = json.loads((ICI / "expediteurs.json").read_text())
legitimes = {k.lower() for k in memoire["legitimes"]}
spam = {k.lower() for k in memoire["spam_confirme_domaines"]}


def classe(adresse, domaine):
    if adresse in legitimes or domaine in legitimes:
        return "legitime"
    return "spam_connu" if domaine in spam else "a_classer"


groupes = defaultdict(lambda: {"nombre": 0, "adresses": set(), "exemples": []})
releve = json.loads((TRAVAIL / sys.argv[1]).read_text())
if isinstance(releve, str):  # browser_evaluate sauve tel quel ce que rend la page : un JSON.stringify arrive doublement encodé
    releve = json.loads(releve)
for m in releve:
    adresse = str(m.get("from") or "").strip().lower()
    domaine = adresse.rpartition("@")[2] or "(sans adresse)"
    g = groupes[(classe(adresse, domaine), domaine)]
    g["nombre"] += 1
    g["adresses"].add(adresse)
    if len(g["exemples"]) < 2:
        g["exemples"].append(str(m.get("l") or "")[:120])

print(json.dumps({"total": len(releve), "groupes": [
    {"classe": c, "domaine": d, "nombre": g["nombre"], "adresses": sorted(g["adresses"])[:5], "exemples": g["exemples"]}
    for (c, d), g in sorted(groupes.items(), key=lambda x: (x[0][0], -x[1]["nombre"]))]}, ensure_ascii=False, indent=0))
