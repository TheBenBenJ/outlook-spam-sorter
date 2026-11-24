#!/opt/homebrew/bin/python3 -I
"""Record an outlook-spam-sorter run in senders.json, approved.json, and outlook-state.json.

Only addresses already received in Gmail enter ``legitimes``, and consumer mail
providers never enter ``spam_confirme_domaines`` so one spam cannot ban all users.

Usage:
    record.py --examined N --restored N --deleted N --to-review N
              [--recovered N --archived N --remaining N]
              [--spam domain ...] [--legitimate address ...] [--approved address ...]
    record.py --failure "reason"

Prints what was learned as JSON for the activity journal:
``{"legitimate": [...], "spam": [...]}``.
Files live in ``$OUTLOOK_SORTER_DATA`` (default: this folder), keeping French JSON
keys intact for the Dashboard and other readers.
"""

import argparse
import datetime
import json
import os
import re
from pathlib import Path
from typing import Any

DATA = Path(os.environ.get("OUTLOOK_SORTER_DATA") or Path(__file__).resolve().parent)
DOMAIN_RE = re.compile(r"[a-z0-9.-]+\.[a-z]{2,}")
HISTORY_LIMIT = 120


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI argument parser for run statistics and learned lists."""
    parser = argparse.ArgumentParser()
    for name in ("--examined", "--restored", "--deleted", "--to-review", "--recovered", "--archived", "--remaining"):
        parser.add_argument(name, type=int, default=0)
    parser.add_argument("--spam", nargs="*", default=[])
    parser.add_argument("--legitimate", nargs="*", default=[])
    parser.add_argument("--approved", nargs="*", default=[])
    parser.add_argument("--failure", default=None)
    return parser


def write_json(path: Path, data: Any) -> None:
    """Write formatted UTF-8 JSON atomically via a sibling temporary file."""
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    os.replace(tmp, path)


def record_failure(data_dir: Path, reason: str, now_iso: str) -> None:
    """Update outlook-state.json with a failure timestamp and truncated reason."""
    state_path = data_dir / "outlook-state.json"
    previous = json.loads(state_path.read_text()) if state_path.exists() else {}
    write_json(state_path, {**previous, "maj": now_iso, "echec": reason[:160]})


def _load_known_gmail(memory: dict[str, Any]) -> dict[str, Any]:
    """Load the optional Gmail known-senders index referenced by senders.json."""
    try:
        return json.loads(Path(memory["gmail_connus"]).read_text())
    except KeyError, OSError, ValueError:
        return {"adresses": {}, "domaines": {}}


def record_run(data_dir: Path, args: argparse.Namespace, now_iso: str) -> dict[str, list[str]]:
    """Persist run metrics, learned senders/domains, approved list, and state file."""
    memory_path = data_dir / "senders.json"
    approved_path = data_dir / "approved.json"
    state_path = data_dir / "outlook-state.json"

    mem = json.loads(memory_path.read_text())
    known = _load_known_gmail(mem)
    approved = set(json.loads(approved_path.read_text())) if approved_path.exists() else set()
    providers = set(known.get("fournisseurs") or [])
    learned: dict[str, list[str]] = {"legitimate": [], "spam": []}
    today_iso = datetime.date.today().isoformat()

    for address in sorted({item.lower() for item in args.legitimate}):
        domain = address.rpartition("@")[2]
        known_domain = bool(providers) and domain not in providers and known["domaines"].get(domain, 0) >= 5
        if address in known["adresses"] or known_domain:
            if address not in mem["legitimes"]:
                learned["legitimate"].append(address)
            # French note matching the rest of senders.json (human-readable).
            mem["legitimes"].setdefault(address, f"{today_iso} déjà reçu dans Gmail")
            approved.add(address)

    approved |= {item.lower() for item in args.approved}
    protected = {key.lower().rpartition("@")[2] for key in mem["legitimes"]}
    additions = {s.lower() for s in args.spam if DOMAIN_RE.fullmatch(s.lower())} - protected - providers
    learned["spam"] = sorted(additions - set(mem["spam_confirme_domaines"]))
    mem["spam_confirme_domaines"] = sorted(set(mem["spam_confirme_domaines"]) | additions)
    mem["historique"].append(
        {
            "date": today_iso,
            "mode": "auto",
            "examines": args.examined,
            "remis_en_boite": args.restored,
            "desabonnements": 0,
            "supprimes": args.deleted,
            "a_verifier": args.to_review,
            "recuperes": args.recovered,
            "archives": args.archived,
        }
    )
    mem["historique"] = mem["historique"][-HISTORY_LIMIT:]

    write_json(memory_path, mem)
    write_json(approved_path, sorted(approved))
    write_json(
        state_path,
        {
            "maj": now_iso,
            "a_verifier": args.to_review,
            "recuperes": args.recovered,
            "archives": args.archived,
            "restants": args.remaining,
            "echec": None,
        },
    )
    return learned


def main(argv: list[str] | None = None) -> None:
    """CLI entry point: record either a failure or a completed sorting pass."""
    args = build_parser().parse_args(argv)
    now_iso = datetime.datetime.now().isoformat(timespec="minutes")
    if args.failure:
        record_failure(DATA, args.failure, now_iso)
        print("failure recorded")
        raise SystemExit(0)
    learned = record_run(DATA, args, now_iso)
    print(json.dumps(learned))


if __name__ == "__main__":
    main()
