#!/opt/homebrew/bin/python3 -I
"""Analyse snapshots of Outlook's mail list, for daily-pass.mjs or imap-pass.py.

Usage:
    analyze.py [--lines] <snapshot.json>   junk folder: groups by domain and class
    analyze.py inbox <snapshot.json>       inbox: recover, archive, review, approve

Snapshots hold ``[{"id", "from", "l"}]`` (``l`` is the row's aria-label, or sender
name and subject from IMAP) and optionally ``"date"`` (YYYY-MM-DD from imap-pass.py),
read only from the ``work/`` folder.
Output keys and class names (``legitime``, ``connu_gmail``, ``spam_connu``,
``a_classer``, ``recuperer``, ``archiver``, ``a_verifier``...) stay French on
purpose: they are the contract shared with senders.json and outlook-state.json.
Each row carries a French ``"motif"`` explaining its decision for the journal.
"""

import datetime
import json
import os
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DATA = Path(os.environ.get("OUTLOOK_SORTER_DATA") or Path(__file__).resolve().parent)
WORK = (DATA / "work").resolve()
MAX_RECOVERIES = 100
MAX_ARCHIVES = 300
RECENT_DAYS = 90
WEEKDAY_ABBR = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
# One-off ticket/order addresses: useless on Outlook's safe senders list.
DISPOSABLE = re.compile(r"\+|[0-9a-f]{8}-[0-9a-f]{4}|^thread-|^[^@]{25,}@")
EMPTY_KNOWN: dict[str, Any] = {"adresses": {}, "domaines": {}, "libelles": {}, "exacts": {}, "recents": []}


@dataclass(slots=True)
class AnalysisContext:
    """Loaded sender rules, Gmail history, and mailbox settings for one pass."""

    legitimate: set[str]
    spam: set[str]
    approved: set[str]
    me: set[str]
    since: str
    noise: tuple[str, ...]
    unreliable: tuple[str, ...]
    known: dict[str, Any]


def load_context(data_dir: Path = DATA) -> AnalysisContext:
    """Load senders.json, optional approved.json, and optional gmail_connus."""
    memory = json.loads((data_dir / "senders.json").read_text())
    try:
        approved = {k.lower() for k in json.loads((data_dir / "approved.json").read_text())}
    except OSError, ValueError:
        approved = set()
    try:
        known = json.loads(Path(memory["gmail_connus"]).read_text())
    except KeyError, OSError, ValueError:
        known = dict(EMPTY_KNOWN)
    return AnalysisContext(
        legitimate={k.lower() for k in memory["legitimes"]},
        spam={k.lower() for k in memory["spam_confirme_domaines"]},
        approved=approved,
        me={k.lower() for k in memory.get("moi", [])},
        since=memory.get("recuperer_depuis", "2000-01-01"),
        noise=tuple(memory.get("libelles_bruit", [])),
        unreliable=tuple(memory.get("libelles_non_fiables", [])),
        known=known,
    )


def known_in_gmail(address: str, domain: str, known: dict[str, Any]) -> bool:
    """Return True if address was seen in Gmail or its non-provider domain appeared >= 5 times."""
    if address in known["adresses"]:
        return True
    providers = known.get("fournisseurs") or []
    return bool(providers) and domain not in providers and known["domaines"].get(domain, 0) >= 5


def class_of(address: str, domain: str, ctx: AnalysisContext) -> tuple[str, str]:
    """Classify a junk-folder sender into ``(class_name, french_motif)``."""
    if address in ctx.legitimate:
        return "legitime", "adresse reconnue comme légitime"
    if domain in ctx.legitimate:
        return "legitime", "domaine reconnu comme légitime"
    if domain in ctx.spam:
        return "spam_connu", "domaine connu pour le spam"
    if known_in_gmail(address, domain, ctx.known):
        return "connu_gmail", "expéditeur déjà reçu dans Gmail"
    return "a_classer", "expéditeur inconnu"


def analyze_junk_rows(
    snapshot: list[dict[str, Any]],
    ctx: AnalysisContext,
    *,
    with_lines: bool = False,
) -> dict[str, Any]:
    """Group junk-folder snapshot rows by ``(class, domain)`` and optionally include row details."""
    groups: dict[tuple[str, str], dict[str, Any]] = defaultdict(
        lambda: {"nombre": 0, "adresses": set(), "exemples": []}
    )
    rows: dict[str, list[dict[str, Any]]] = {"legitime": [], "connu_gmail": [], "spam_connu": [], "a_classer": []}
    for item in snapshot:
        address = str(item.get("from") or "").strip().lower()
        domain = address.rpartition("@")[2] or "(sans adresse)"
        cls, why = class_of(address, domain, ctx)
        group = groups[(cls, domain)]
        group["nombre"] += 1
        group["adresses"].add(address)
        if len(group["exemples"]) < 2:
            group["exemples"].append(str(item.get("l") or "")[:120])
        rows[cls].append({"id": item["id"], "from": address, "objet": str(item.get("l") or "")[:160], "motif": why})
    output: dict[str, Any] = {
        "total": len(snapshot),
        "groupes": [
            {
                "classe": cls,
                "domaine": domain,
                "nombre": group["nombre"],
                "adresses": sorted(group["adresses"])[:5],
                "exemples": group["exemples"],
            }
            for (cls, domain), group in sorted(groups.items(), key=lambda x: (x[0][0], -x[1]["nombre"]))
        ],
    }
    if with_lines:
        output["lignes"] = rows
    return output


def flatten(text: str | None) -> str:
    """Normalize text to lowercase single-spaced NFC UTF-8 for fuzzy subject matching."""
    cleaned = unicodedata.normalize("NFC", (text or "").encode("utf-8", "ignore").decode("utf-8"))
    return re.sub(r"\s+", " ", cleaned.lower().replace("\xa0", " ")).strip()


def row_date(row_text: str, today: datetime.date | None = None) -> datetime.date | None:
    """Parse a date from Outlook's French UI row label (DD/MM/YYYY, weekday DD/MM, or time)."""
    today = today or datetime.date.today()
    match = re.search(r"\b(\d{2})/(\d{2})/(20\d{2})\b", row_text)
    if match:
        return datetime.date(int(match[3]), int(match[2]), int(match[1]))
    match = re.search(r"\b(?:Lun|Mar|Mer|Jeu|Ven|Sam|Dim) (\d{2})/(\d{2})\b", row_text)
    if match:
        parsed = datetime.date(today.year, int(match[2]), int(match[1]))
        return parsed if parsed <= today else parsed.replace(year=today.year - 1)
    match = re.search(r"\b(Lun|Mar|Mer|Jeu|Ven|Sam|Dim) \d{1,2}:\d{2}\b", row_text)
    if match:
        target_weekday = WEEKDAY_ABBR.index(match[1])
        for offset in range(8):
            candidate = today - datetime.timedelta(days=offset)
            if candidate.weekday() == target_weekday:
                return candidate
    if re.search(r"(?<![à:] )\b\d{1,2}:\d{2}\b", row_text) and not re.search(r"réunion", row_text, re.I):
        return today
    return None


def gmail_label(address: str, known: dict[str, Any]) -> str | None:
    """Return the Gmail label recorded for an exact address or its domain suffix."""
    if address in known["exacts"]:
        return known["exacts"][address]
    parts = address.rpartition("@")[2].split(".")
    for index in range(len(parts) - 1):
        label = known["libelles"].get(".".join(parts[index:]))
        if label:
            return label
    return None


def trust_reason(address: str, domain: str, label: str, ctx: AnalysisContext) -> str | None:
    """Return why a sender is trusted for inbox recovery, or None if unknown/untrusted."""
    if address in ctx.legitimate:
        return "adresse reconnue comme légitime"
    if domain in ctx.legitimate:
        return "domaine reconnu comme légitime"
    if known_in_gmail(address, domain, ctx.known):
        return "expéditeur déjà reçu dans Gmail"
    if label and not label.startswith(ctx.unreliable):
        return f"expéditeur rangé dans Gmail sous « {label} »"
    return None


@dataclass(slots=True)
class InboxCandidate:
    """Normalized metadata of one inbox mail used to pick its action bucket."""

    address: str
    domain: str
    label: str
    day: str
    has_date: bool
    already_in_gmail: bool
    recent_since: str


def _classify_inbox_row(
    item: dict[str, Any],
    ctx: AnalysisContext,
    recent_mails: dict[str, list[tuple[str, str]]],
    recent_since: str,
) -> tuple[str, dict[str, Any]]:
    """Assign an inbox row to ``'recover'``, ``'archive'``, or ``'review'`` with its French motif."""
    address = str(item.get("from") or "").strip().lower()
    row_text = str(item.get("l") or "")
    parsed_date = datetime.date.fromisoformat(item["date"]) if item.get("date") else row_date(row_text)
    # Unreadable dates are treated as old so the mail is archived instead of stuck forever.
    day = parsed_date.isoformat() if parsed_date else "0000-00-00"
    flat_text = flatten(row_text)
    window = {(parsed_date + datetime.timedelta(days=k)).isoformat() for k in (-1, 0, 1)} if parsed_date else set()
    already_in_gmail = any(d in window and subj[:30] in flat_text for d, subj in recent_mails.get(address, []) if subj)
    label = gmail_label(address, ctx.known) or ""
    domain = address.rpartition("@")[2]
    row: dict[str, Any] = {"id": item["id"], "from": address, "date": day, "objet": row_text[:160]}
    candidate = InboxCandidate(
        address=address,
        domain=domain,
        label=label,
        day=day,
        has_date=parsed_date is not None,
        already_in_gmail=already_in_gmail,
        recent_since=recent_since,
    )
    bucket, motif = _inbox_bucket_and_motif(candidate, ctx)
    row["motif"] = motif
    return bucket, row


def _inbox_bucket_and_motif(candidate: InboxCandidate, ctx: AnalysisContext) -> tuple[str, str]:
    """Return ``(bucket, french_motif)`` for one inbox message."""
    if candidate.address in ctx.me:
        return "archive", "envoyé depuis une de mes adresses"
    if candidate.already_in_gmail:
        return "archive", "déjà dans Gmail"
    if candidate.day < ctx.since:
        return "archive", f"reçu avant le {ctx.since}" if candidate.has_date else "date de réception illisible"
    if ctx.noise and candidate.label.startswith(ctx.noise):
        return "archive", f"bruit, rangé dans Gmail sous « {candidate.label} »"
    if candidate.domain in ctx.spam:
        return "review", "domaine connu pour le spam"
    trust = trust_reason(candidate.address, candidate.domain, candidate.label, ctx)
    if not trust:
        # Never recover or approve an unknown sender: phishing can slip past Outlook filters.
        if candidate.day >= candidate.recent_since:
            return "review", "expéditeur inconnu"
        return "archive", f"expéditeur inconnu, mail de plus de {RECENT_DAYS} jours"
    return "recover", trust


def analyze_inbox_rows(
    snapshot: list[dict[str, Any]],
    ctx: AnalysisContext,
    *,
    today: datetime.date | None = None,
) -> dict[str, Any]:
    """Split inbox rows into bounded recover, archive, review, and approve lists."""
    recent_mails: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for sender, subject, day in ctx.known["recents"]:
        recent_mails[sender].append((day, subject))
    buckets: dict[str, list[dict[str, Any]]] = {"recover": [], "archive": [], "review": []}
    ref_day = today or datetime.date.today()
    recent_since = (ref_day - datetime.timedelta(days=RECENT_DAYS)).isoformat()
    for item in snapshot:
        bucket, row = _classify_inbox_row(item, ctx, recent_mails, recent_since)
        buckets[bucket].append(row)
    for rows in buckets.values():
        rows.sort(key=lambda x: x["date"], reverse=True)
    recover, archive, review = buckets["recover"], buckets["archive"], buckets["review"]
    batch = recover[:MAX_RECOVERIES]
    approve = sorted({x["from"] for x in batch if x["from"] not in ctx.approved and not DISPOSABLE.search(x["from"])})
    return {
        "total": len(snapshot),
        "a_verifier": review[:50],
        "a_verifier_total": len(review),
        "recuperer": batch,
        "recuperer_restants": len(recover) - len(batch),
        "archiver": archive[:MAX_ARCHIVES],
        "archiver_restants": max(0, len(archive) - MAX_ARCHIVES),
        "approuver": approve,
    }


def _load_snapshot(path: Path) -> list[dict[str, Any]]:
    """Read a JSON snapshot file, unwrapping browser_evaluate double-encoded JSON if needed."""
    data = json.loads(path.read_text())
    if isinstance(data, str):
        data = json.loads(data)
    return data


def main(argv: list[str] | None = None) -> None:
    """CLI entry point: parse arguments, analyse snapshot from work/, and print JSON."""
    args = list(sys.argv[1:] if argv is None else argv)
    mode = args.pop(0) if args and args[0] == "inbox" else "junk"
    with_lines = "--lines" in args
    if with_lines:
        args.remove("--lines")
    if len(args) != 1 or (WORK / args[0]).resolve().parent != WORK:
        sys.exit("usage: analyze.py [inbox | --lines] <file in the work/ folder>")
    ctx = load_context(DATA)
    snap = _load_snapshot(WORK / args[0])
    result = analyze_inbox_rows(snap, ctx) if mode == "inbox" else analyze_junk_rows(snap, ctx, with_lines=with_lines)
    print(json.dumps(result, ensure_ascii=False, indent=0))


if __name__ == "__main__":
    main()
