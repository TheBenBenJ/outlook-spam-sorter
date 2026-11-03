#!/opt/homebrew/bin/python3 -I
"""Analyse les relevés d'Outlook, pour daily-pass.mjs ou une séance interactive.

analyze.py [--lignes] <relevé.json>   dossier indésirable : groupes par domaine et par classe ; --lignes y ajoute les lignes de chaque classe
analyze.py boite <relevé.json>        boîte de réception : quoi récupérer dans Gmail, quoi archiver, qui ajouter à la liste blanche

Relevés : liste de {"id", "from", "l"}, dans le dossier work/ uniquement.
"""
import datetime, json, re, sys, unicodedata
from collections import defaultdict
from pathlib import Path

ICI = Path(__file__).resolve().parent
TRAVAIL = (ICI / "work").resolve()
MAX_RECUPERATIONS, MAX_ARCHIVES = 100, 300  # par passage, pour en borner la durée : le reste attend le passage suivant
RECENT_JOURS = 90  # un inconnu plus ancien part aux archives sans récupération : il n'a plus rien d'urgent
JETABLE = re.compile(r"\+|[0-9a-f]{8}-[0-9a-f]{4}|^thread-|^[^@]{25,}@")  # adresse propre à un ticket ou à une commande : inutile sur la liste blanche

args = sys.argv[1:]
mode = args.pop(0) if args and args[0] == "boite" else "indesirable"
avec_lignes = "--lignes" in args
if avec_lignes:
    args.remove("--lignes")
if len(args) != 1 or (TRAVAIL / args[0]).resolve().parent != TRAVAIL:
    sys.exit("usage : analyze.py [boite | --lignes] <fichier du dossier work/>")
memoire = json.loads((ICI / "expediteurs.json").read_text())
legitimes = {k.lower() for k in memoire["legitimes"]}
spam = {k.lower() for k in memoire["spam_confirme_domaines"]}
try:
    approuves = {k.lower() for k in json.loads((ICI / "approuves.json").read_text())}  # liste blanche relevée dans Outlook : à part, l'agent n'a pas à la lire
except (OSError, ValueError):
    approuves = set()
moi = {k.lower() for k in memoire.get("moi", [])}
depuis = memoire.get("recuperer_depuis", "2000-01-01")
bruit = tuple(memoire.get("libelles_bruit", []))
non_fiables = tuple(memoire.get("libelles_non_fiables", []))
try:
    connus = json.loads(Path(memoire["gmail_connus"]).read_text())
except (KeyError, OSError, ValueError):
    connus = {"adresses": {}, "domaines": {}, "libelles": {}, "exacts": {}, "recents": []}


def connu_gmail(adresse, domaine):
    """Déjà reçu dans Gmail : l'adresse exacte, ou un domaine d'entreprise vu au moins cinq fois. Sans liste des messageries grand public, le domaine ne prouve rien."""
    if adresse in connus["adresses"]:
        return True
    fournisseurs = connus.get("fournisseurs") or []
    return bool(fournisseurs) and domaine not in fournisseurs and connus["domaines"].get(domaine, 0) >= 5


def classe(adresse, domaine):
    if adresse in legitimes or domaine in legitimes:
        return "legitime"
    if domaine in spam:
        return "spam_connu"
    return "connu_gmail" if connu_gmail(adresse, domaine) else "a_classer"


releve = json.loads((TRAVAIL / args[0]).read_text())
if isinstance(releve, str):  # browser_evaluate sauve tel quel ce que rend la page : un JSON.stringify arrive doublement encodé
    releve = json.loads(releve)


def analyse_indesirable():
    groupes = defaultdict(lambda: {"nombre": 0, "adresses": set(), "exemples": []})
    lignes = {"legitime": [], "connu_gmail": [], "spam_connu": [], "a_classer": []}
    for m in releve:
        adresse = str(m.get("from") or "").strip().lower()
        domaine = adresse.rpartition("@")[2] or "(sans adresse)"
        c = classe(adresse, domaine)
        g = groupes[(c, domaine)]
        g["nombre"] += 1
        g["adresses"].add(adresse)
        if len(g["exemples"]) < 2:
            g["exemples"].append(str(m.get("l") or "")[:120])
        lignes[c].append({"id": m["id"], "from": adresse, "objet": str(m.get("l") or "")[:160]})
    sortie = {"total": len(releve), "groupes": [
        {"classe": c, "domaine": d, "nombre": g["nombre"], "adresses": sorted(g["adresses"])[:5], "exemples": g["exemples"]}
        for (c, d), g in sorted(groupes.items(), key=lambda x: (x[0][0], -x[1]["nombre"]))]}
    if avec_lignes:
        sortie["lignes"] = lignes
    print(json.dumps(sortie, ensure_ascii=False, indent=0))


def plat(t):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", (t or "").encode("utf-8", "ignore").decode("utf-8")).lower().replace("\xa0", " ")).strip()


def jour_ligne(l):
    """Date lue dans le libellé de la liste : « 29/09/2026 », « Mar 27/10 », « Mer 17:45 » ou « 17:45 ». None si elle n'y est pas (invitation d'agenda, objet très long)."""
    auj = datetime.date.today()
    m = re.search(r"\b(\d{2})/(\d{2})/(20\d{2})\b", l)
    if m:
        return datetime.date(int(m[3]), int(m[2]), int(m[1]))
    m = re.search(r"\b(?:Lun|Mar|Mer|Jeu|Ven|Sam|Dim) (\d{2})/(\d{2})\b", l)
    if m:
        d = datetime.date(auj.year, int(m[2]), int(m[1]))
        return d if d <= auj else d.replace(year=auj.year - 1)
    m = re.search(r"\b(Lun|Mar|Mer|Jeu|Ven|Sam|Dim) \d{1,2}:\d{2}\b", l)
    if m:
        for k in range(8):
            d = auj - datetime.timedelta(days=k)
            if d.weekday() == ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"].index(m[1]):
                return d
    if re.search(r"(?<![à:] )\b\d{1,2}:\d{2}\b", l) and not re.search(r"réunion", l, re.I):
        return auj
    return None


def libelle(adresse):
    if adresse in connus["exacts"]:
        return connus["exacts"][adresse]
    parties = adresse.rpartition("@")[2].split(".")
    for i in range(len(parties) - 1):
        lib = connus["libelles"].get(".".join(parties[i:]))
        if lib:
            return lib
    return None


def analyse_boite():
    recents = defaultdict(list)
    for de, objet, jour in connus["recents"]:
        recents[de].append((jour, objet))
    recuperer, archiver, a_verifier = [], [], []
    recent = (datetime.date.today() - datetime.timedelta(days=RECENT_JOURS)).isoformat()
    for m in releve:
        adresse = str(m.get("from") or "").strip().lower()
        l = str(m.get("l") or "")
        d = jour_ligne(l)
        jour = d.isoformat() if d else "0000-00-00"  # date illisible : traité comme ancien, le mail reste dans Archivage
        texte = plat(l)
        marge = {(d + datetime.timedelta(days=k)).isoformat() for k in (-1, 0, 1)} if d else set()
        deja = any(j in marge and o[:30] in texte for j, o in recents.get(adresse, []) if o)
        lib = libelle(adresse) or ""
        ligne = {"id": m["id"], "from": adresse, "date": jour, "objet": l[:160]}
        domaine = adresse.rpartition("@")[2]
        fiable = adresse in legitimes or domaine in legitimes or connu_gmail(adresse, domaine) or bool(lib and not lib.startswith(non_fiables))
        if adresse in moi or deja or jour < depuis or (bruit and lib.startswith(bruit)):
            archiver.append(ligne)
        elif domaine in spam:
            a_verifier.append(ligne)
        elif not fiable:  # la boîte contient aussi de l'hameçonnage passé à travers les filtres : jamais de récupération ni de liste blanche pour un inconnu
            (a_verifier if jour >= recent else archiver).append(ligne)
        else:
            recuperer.append(ligne)
    recuperer.sort(key=lambda x: x["date"], reverse=True)
    archiver.sort(key=lambda x: x["date"], reverse=True)
    lot = recuperer[:MAX_RECUPERATIONS]
    approuver = sorted({x["from"] for x in lot if x["from"] not in approuves and not JETABLE.search(x["from"])})
    a_verifier.sort(key=lambda x: x["date"], reverse=True)
    print(json.dumps({"total": len(releve), "a_verifier": a_verifier[:50], "a_verifier_total": len(a_verifier), "recuperer": lot, "recuperer_restants": len(recuperer) - len(lot),
                      "archiver": archiver[:MAX_ARCHIVES], "archiver_restants": max(0, len(archiver) - MAX_ARCHIVES), "approuver": approuver}, ensure_ascii=False, indent=0))


analyse_boite() if mode == "boite" else analyse_indesirable()
