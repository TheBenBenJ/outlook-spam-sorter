# AGENTS.md — outlook-spam-sorter

Quick reference for coding agents working in this repository. Detailed engineering rules live in [`CLAUDE.md`](CLAUDE.md); full system documentation lives in [`README.md`](README.md).

## Layout

| Path | Role |
|---|---|
| `imap-pass.py` | Main server-side IMAP pass (junk cleanup, stuck-inbox recovery to Gmail, safe-sender tracking). |
| `outlook_imap.py` | IMAP + OAuth2 client wrapper for Outlook and Gmail. |
| `analyze.py` | Deterministic sender/rule evaluation for junk and inbox snapshots. |
| `classify.py` | Tool-less local LLM classifier (strict JSON schema, confidence threshold). |
| `record.py` | Updates `senders.json`, `approved.json`, and `outlook-state.json` (read by Dashboard). |
| `activity.py` | Mode-600 JSONL activity journal with 90-day / size rotation. |
| `daily-pass.mjs` | macOS headless-Chrome fallback pass (launched by `run.sh`). |
| `tests.py` | Deterministic offline test suite (fake IMAP + fake local model, `python3 -I tests.py`). |
| `skill/` | Claude Code interactive review skill for doubtful mails. |

## Verification Commands

```bash
ruff check . && ruff format --check .
python3 -I tests.py
node --check daily-pass.mjs
zsh -n run.sh
```

## Deployment

GitHub Actions (`.github/workflows/ci.yml`) runs `py_compile`, Ruff lint & format, `python -I tests.py`, `node --check`, `bash -n`/`shellcheck`, `zsh -n`, and JSON checks. A green push on `main` fast-forwards `production`, which `deploy-watch` applies on `pc-benjamin`.
