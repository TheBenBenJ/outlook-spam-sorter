#!/usr/bin/env python3
"""Deterministic offline unit tests for pure logic in analyze.py, record.py, classify.py, and outlook_imap.py.

Run offline with isolated Python:
    python3 -I tests/test_outlook_sorter.py
"""

import datetime
import importlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

analyze = importlib.import_module("analyze")
classify = importlib.import_module("classify")
outlook_imap = importlib.import_module("outlook_imap")
record = importlib.import_module("record")


def _sample_context() -> analyze.AnalysisContext:
    """Return a deterministic in-memory AnalysisContext for unit tests."""
    return analyze.AnalysisContext(
        legitimate={"trusted@bank.example", "family.example"},
        spam={"casino.example"},
        approved={"already@approved.example"},
        me={"me@hotmail.example"},
        since="2026-01-01",
        noise=("Pub/",),
        unreliable=("Indésirables/",),
        known={
            "adresses": {"friend@gmail.com": 3},
            "domaines": {"acme.example": 7, "gmail.com": 50},
            "fournisseurs": ["gmail.com"],
            "libelles": {"shop.example": "Pub/Promos", "utility.example": "Factures"},
            "exacts": {"alert@service.example": "Indésirables/Alertes"},
            "recents": [["friend@gmail.com", "photos du weekend", "2026-10-15"]],
        },
    )


class AnalyzeUnitTests(unittest.TestCase):
    """Pure rule tests for analyze.py junk and inbox classification."""

    def test_known_in_gmail_and_class_of(self) -> None:
        ctx = _sample_context()
        self.assertTrue(analyze.known_in_gmail("friend@gmail.com", "gmail.com", ctx.known))
        self.assertFalse(analyze.known_in_gmail("stranger@gmail.com", "gmail.com", ctx.known))
        self.assertTrue(analyze.known_in_gmail("hr@acme.example", "acme.example", ctx.known))
        self.assertEqual(analyze.class_of("trusted@bank.example", "bank.example", ctx)[0], "legitime")
        self.assertEqual(analyze.class_of("mom@family.example", "family.example", ctx)[0], "legitime")
        self.assertEqual(analyze.class_of("win@casino.example", "casino.example", ctx)[0], "spam_connu")
        self.assertEqual(analyze.class_of("hr@acme.example", "acme.example", ctx)[0], "connu_gmail")
        self.assertEqual(analyze.class_of("unknown@random.example", "random.example", ctx)[0], "a_classer")

    def test_row_date_parsing(self) -> None:
        today = datetime.date(2026, 11, 4)  # Wednesday
        self.assertEqual(analyze.row_date("Banque 29/09/2026 Relevé", today), datetime.date(2026, 9, 29))
        self.assertEqual(analyze.row_date("Newsletter Mar 27/10", today), datetime.date(2026, 10, 27))
        self.assertEqual(analyze.row_date("Facture Sam 20/12", today), datetime.date(2025, 12, 20))
        self.assertEqual(analyze.row_date("Ami Lun 14:30", today), datetime.date(2026, 11, 2))
        self.assertEqual(analyze.row_date("Banque 17:45", today), today)
        self.assertIsNone(analyze.row_date("Invitation réunion 14:00", today))

    def test_analyze_inbox_buckets_and_disposable_filter(self) -> None:
        ctx = _sample_context()
        today = datetime.date(2026, 11, 4)
        snapshot = [
            {"id": "1", "from": "me@hotmail.example", "l": "Note", "date": "2026-11-01"},
            {"id": "2", "from": "friend@gmail.com", "l": "Photos du weekend dernier", "date": "2026-10-16"},
            {"id": "3", "from": "billing@utility.example", "l": "Facture", "date": "2026-11-02"},
            {"id": "4", "from": "promo@shop.example", "l": "Soldes", "date": "2026-11-02"},
            {"id": "5", "from": "ticket+123@acme.example", "l": "Ticket", "date": "2026-11-03"},
            {"id": "6", "from": "alert@service.example", "l": "Alerte", "date": "2026-11-03"},
            {"id": "7", "from": "old@unknown.example", "l": "Vieux", "date": "2026-02-01"},
            {"id": "8", "from": "win@casino.example", "l": "Jackpot", "date": "2026-11-03"},
        ]
        result = analyze.analyze_inbox_rows(snapshot, ctx, today=today)
        self.assertEqual([r["id"] for r in result["recuperer"]], ["5", "3"])
        self.assertEqual(result["approuver"], ["billing@utility.example"])
        self.assertEqual({r["id"] for r in result["archiver"]}, {"1", "2", "4", "7"})
        self.assertEqual({r["id"] for r in result["a_verifier"]}, {"6", "8"})


class RecordAndClassifyUnitTests(unittest.TestCase):
    """Pure rule tests for record.py, classify.py, and outlook_imap.py."""

    def test_record_protects_legitimate_and_providers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data = Path(directory)
            known = {"adresses": {"ok@gmail.com": 2}, "domaines": {}, "fournisseurs": ["gmail.com"]}
            (data / "known.json").write_text(json.dumps(known))
            senders = {
                "legitimes": {"safe@keep.example": "ok"},
                "spam_confirme_domaines": [],
                "historique": [],
                "gmail_connus": str(data / "known.json"),
            }
            (data / "senders.json").write_text(json.dumps(senders))
            args = record.build_parser().parse_args(
                [
                    "--examined",
                    "10",
                    "--restored",
                    "1",
                    "--deleted",
                    "3",
                    "--to-review",
                    "0",
                    "--spam",
                    "gmail.com",
                    "keep.example",
                    "bad.example",
                    "not-a-domain",
                    "--legitimate",
                    "ok@gmail.com",
                    "unknown@nope.example",
                ]
            )
            learned = record.record_run(data, args, "2026-11-04T07:30")
            self.assertEqual(learned, {"legitimate": ["ok@gmail.com"], "spam": ["bad.example"]})
            saved = json.loads((data / "senders.json").read_text())
            self.assertEqual(saved["spam_confirme_domaines"], ["bad.example"])
            self.assertIn("ok@gmail.com", saved["legitimes"])
            self.assertNotIn("unknown@nope.example", saved["legitimes"])

    def test_classify_keep_filters_invalid_and_conflicting(self) -> None:
        answers = [
            {"n": 0, "verdict": "legitime", "confiance": 0.9},
            {"n": 1, "verdict": "spam", "confiance": 1.2},
            {"n": 2, "verdict": "spam", "confiance": 0.8},
            {"n": 2, "verdict": "legitime", "confiance": 0.7},
            {"n": 99, "verdict": "spam", "confiance": 1.0},
            {"n": True, "verdict": "spam", "confiance": 1.0},
        ]
        kept = classify.keep(answers, {0, 1, 2})
        self.assertEqual(kept, {"0": {"verdict": "legitime", "confiance": 0.9}})

    def test_mutf7_decoding(self) -> None:
        self.assertEqual(outlook_imap.mutf7("Courrier ind&AOk-sirable"), "Courrier indésirable")
        self.assertEqual(outlook_imap.mutf7("&AMk-l&AOk-ments supprim&AOk-s"), "Éléments supprimés")
        self.assertEqual(outlook_imap.mutf7("A&-B"), "A&B")


if __name__ == "__main__":
    unittest.main(verbosity=2)
