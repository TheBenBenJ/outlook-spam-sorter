#!/opt/homebrew/bin/python3 -I
"""Analyse snapshots of Outlook's mail list, for daily-pass.mjs or an interactive session.

analyze.py [--lines] <snapshot.json>   junk folder: groups by domain and class; --lines adds the rows of each class
analyze.py inbox <snapshot.json>       inbox: what to recover into Gmail, what to archive, whom to add to the safe senders list

Snapshots: a list of {"id", "from", "l"} ("l" is the row's aria-label, or sender name and subject from IMAP) and, from imap-pass.py,
"date" (YYYY-MM-DD, the received date), read from the work/ folder only.
Output keys and class names (legitime, connu_gmail, spam_connu, a_classer, recuperer, archiver, a_verifier...) stay French on
purpose: they are the vocabulary of senders.json and outlook-state.json, which other components read (see the README).
Each row carries a "motif": why it falls in its class or list, in French, for the activity journal that Benjamin reads.
"""
import datetime, json, re, sys, unicodedata
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
WORK = (HERE / "work").resolve()
MAX_RECOVERIES, MAX_ARCHIVES = 100, 300  # per run, to bound its duration: the rest waits for the next run
RECENT_DAYS = 90  # an older unknown sender is archived without recovery: nothing urgent is left in it
DISPOSABLE = re.compile(r"\+|[0-9a-f]{8}-[0-9a-f]{4}|^thread-|^[^@]{25,}@")  # address specific to one ticket or order: useless on the safe senders list

args = sys.argv[1:]
mode = args.pop(0) if args and args[0] == "inbox" else "junk"
with_lines = "--lines" in args
if with_lines:
    args.remove("--lines")
if len(args) != 1 or (WORK / args[0]).resolve().parent != WORK:
    sys.exit("usage: analyze.py [inbox | --lines] <file in the work/ folder>")
memory = json.loads((HERE / "senders.json").read_text())
legitimate = {k.lower() for k in memory["legitimes"]}
spam = {k.lower() for k in memory["spam_confirme_domaines"]}
try:
    approved = {k.lower() for k in json.loads((HERE / "approved.json").read_text())}  # safe senders list read from Outlook: kept apart, the agent does not need to read it
except (OSError, ValueError):
    approved = set()
me = {k.lower() for k in memory.get("moi", [])}
since = memory.get("recuperer_depuis", "2000-01-01")
noise = tuple(memory.get("libelles_bruit", []))
unreliable = tuple(memory.get("libelles_non_fiables", []))
try:
    known = json.loads(Path(memory["gmail_connus"]).read_text())
except (KeyError, OSError, ValueError):
    known = {"adresses": {}, "domaines": {}, "libelles": {}, "exacts": {}, "recents": []}


def known_in_gmail(address, domain):
    """Already received in Gmail: the exact address, or a company domain seen at least five times. Without the list of consumer mail providers, a domain proves nothing."""
    if address in known["adresses"]:
        return True
    providers = known.get("fournisseurs") or []
    return bool(providers) and domain not in providers and known["domaines"].get(domain, 0) >= 5


def class_of(address, domain):
    """-> (class, motif)."""
    if address in legitimate or domain in legitimate:
        return "legitime", "adresse reconnue comme légitime" if address in legitimate else "domaine reconnu comme légitime"
    if domain in spam:
        return "spam_connu", "domaine connu pour le spam"
    return ("connu_gmail", "expéditeur déjà reçu dans Gmail") if known_in_gmail(address, domain) else ("a_classer", "expéditeur inconnu")


snapshot = json.loads((WORK / args[0]).read_text())
if isinstance(snapshot, str):  # browser_evaluate saves what the page returns as is: a JSON.stringify arrives double-encoded
    snapshot = json.loads(snapshot)


def analyze_junk():
    groups = defaultdict(lambda: {"nombre": 0, "adresses": set(), "exemples": []})
    rows = {"legitime": [], "connu_gmail": [], "spam_connu": [], "a_classer": []}
    for m in snapshot:
        address = str(m.get("from") or "").strip().lower()
        domain = address.rpartition("@")[2] or "(sans adresse)"
        c, why = class_of(address, domain)
        g = groups[(c, domain)]
        g["nombre"] += 1
        g["adresses"].add(address)
        if len(g["exemples"]) < 2:
            g["exemples"].append(str(m.get("l") or "")[:120])
        rows[c].append({"id": m["id"], "from": address, "objet": str(m.get("l") or "")[:160], "motif": why})
    output = {"total": len(snapshot), "groupes": [
        {"classe": c, "domaine": d, "nombre": g["nombre"], "adresses": sorted(g["adresses"])[:5], "exemples": g["exemples"]}
        for (c, d), g in sorted(groups.items(), key=lambda x: (x[0][0], -x[1]["nombre"]))]}
    if with_lines:
        output["lignes"] = rows
    print(json.dumps(output, ensure_ascii=False, indent=0))


def flatten(t):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", (t or "").encode("utf-8", "ignore").decode("utf-8")).lower().replace("\xa0", " ")).strip()


def row_date(row_text):
    """Date read from the row's label, as Outlook's French UI writes it: "29/09/2026", "Mar 27/10", "Mer 17:45" or "17:45". None when absent (calendar invitation, very long subject)."""
    today = datetime.date.today()
    m = re.search(r"\b(\d{2})/(\d{2})/(20\d{2})\b", row_text)
    if m:
        return datetime.date(int(m[3]), int(m[2]), int(m[1]))
    m = re.search(r"\b(?:Lun|Mar|Mer|Jeu|Ven|Sam|Dim) (\d{2})/(\d{2})\b", row_text)
    if m:
        d = datetime.date(today.year, int(m[2]), int(m[1]))
        return d if d <= today else d.replace(year=today.year - 1)
    m = re.search(r"\b(Lun|Mar|Mer|Jeu|Ven|Sam|Dim) \d{1,2}:\d{2}\b", row_text)
    if m:
        for k in range(8):
            d = today - datetime.timedelta(days=k)
            if d.weekday() == ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"].index(m[1]):
                return d
    if re.search(r"(?<![à:] )\b\d{1,2}:\d{2}\b", row_text) and not re.search(r"réunion", row_text, re.I):
        return today
    return None


def gmail_label(address):
    if address in known["exacts"]:
        return known["exacts"][address]
    parts = address.rpartition("@")[2].split(".")
    for i in range(len(parts) - 1):
        label = known["libelles"].get(".".join(parts[i:]))
        if label:
            return label
    return None


def analyze_inbox():
    recent_mails = defaultdict(list)
    for sender, subject, day in known["recents"]:
        recent_mails[sender].append((day, subject))
    recover, archive, review = [], [], []
    recent_since = (datetime.date.today() - datetime.timedelta(days=RECENT_DAYS)).isoformat()
    for m in snapshot:
        address = str(m.get("from") or "").strip().lower()
        row_text = str(m.get("l") or "")
        d = datetime.date.fromisoformat(m["date"]) if m.get("date") else row_date(row_text)  # imap-pass.py gives the received date; the browser pass reads it in the row
        day = d.isoformat() if d else "0000-00-00"  # unreadable date: treated as old, the mail stays in the archive
        text = flatten(row_text)
        window = {(d + datetime.timedelta(days=k)).isoformat() for k in (-1, 0, 1)} if d else set()
        already_in_gmail = any(j in window and o[:30] in text for j, o in recent_mails.get(address, []) if o)
        label = gmail_label(address) or ""
        row = {"id": m["id"], "from": address, "date": day, "objet": row_text[:160]}
        domain = address.rpartition("@")[2]
        trust = ("adresse reconnue comme légitime" if address in legitimate else "domaine reconnu comme légitime" if domain in legitimate
                 else "expéditeur déjà reçu dans Gmail" if known_in_gmail(address, domain)
                 else f"expéditeur rangé dans Gmail sous « {label} »" if label and not label.startswith(unreliable) else None)
        if address in me:
            to, row["motif"] = archive, "envoyé depuis une de mes adresses"
        elif already_in_gmail:
            to, row["motif"] = archive, "déjà dans Gmail"
        elif day < since:
            to, row["motif"] = archive, f"reçu avant le {since}" if d else "date de réception illisible"
        elif noise and label.startswith(noise):
            to, row["motif"] = archive, f"bruit, rangé dans Gmail sous « {label} »"
        elif domain in spam:
            to, row["motif"] = review, "domaine connu pour le spam"
        elif not trust:  # the inbox also holds phishing that got through the filters: never recover nor approve an unknown sender
            to, row["motif"] = (review, "expéditeur inconnu") if day >= recent_since else (archive, f"expéditeur inconnu, mail de plus de {RECENT_DAYS} jours")
        else:
            to, row["motif"] = recover, trust
        to.append(row)
    recover.sort(key=lambda x: x["date"], reverse=True)
    archive.sort(key=lambda x: x["date"], reverse=True)
    batch = recover[:MAX_RECOVERIES]
    approve = sorted({x["from"] for x in batch if x["from"] not in approved and not DISPOSABLE.search(x["from"])})
    review.sort(key=lambda x: x["date"], reverse=True)
    print(json.dumps({"total": len(snapshot), "a_verifier": review[:50], "a_verifier_total": len(review), "recuperer": batch, "recuperer_restants": len(recover) - len(batch),
                      "archiver": archive[:MAX_ARCHIVES], "archiver_restants": max(0, len(archive) - MAX_ARCHIVES), "approuver": approve}, ensure_ascii=False, indent=0))


analyze_inbox() if mode == "inbox" else analyze_junk()
