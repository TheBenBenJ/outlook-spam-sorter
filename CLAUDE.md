# outlook-spam-sorter

Daily unattended Outlook.com (Hotmail) junk cleaner and stuck-inbox recovery tool (`imap-pass.py` on `pc-benjamin`, `daily-pass.mjs` browser fallback on Mac). Full architecture & operations: [`README.md`](README.md). Agent quick map: [`AGENTS.md`](AGENTS.md).

## Engineering Standards (Mandatory)

1. **Pragmatic Clean Architecture & Tactical DDD**:
   - Keep external I/O adapters (`outlook_imap.py`, `daily-pass.mjs`, `classify.py`) separated from pure decision/classification logic (`analyze.py`, `record.py`, `activity.py`).
   - Keep pure domain logic independent of IMAP/network/filesystem so it is directly unit-testable offline with `python3 -I`.
   - Stay pragmatic: avoid needless boilerplate (no DI containers or redundant DTO layers on concise scripts).
2. **Typed domain models — no ad-hoc JSON construction**:
   - Model structured domain entities (verdicts, classification results, run summaries, journal events) with stdlib `@dataclass(slots=True)` and explicit serialization helpers instead of building large nested raw dicts inline.
3. **Modular, well-scoped code**:
   - Keep functions focused and concise; standard library only for Python (`python3 -I` compatible).
4. **100% English technical code**:
   - All code, filenames, classes, functions, variables, docstrings, comments, tests, documentation, and commit messages MUST be in **English**.
   - Keep user-facing notifications, Outlook folder names, and frozen Dashboard wire-JSON keys (`outlook-state.json`, `activity.jsonl`) in **French**.
5. **Tests & documentation required**:
   - Every new rule, classifier change, or bug fix MUST include deterministic offline unit tests in `tests.py` (fake IMAP server, fake local model, zero network/secrets).
   - Keep `README.md`, `CLAUDE.md`, and `AGENTS.md` accurate and up to date whenever structure, scripts, or workflows change.
6. **Linting & validation before every commit**:
   - Python: `ruff check . && ruff format --check .` and `python3 -m py_compile *.py`
   - JS/Shell: `node --check daily-pass.mjs`, `zsh -n run.sh`
   - Tests: `python3 -I tests.py`

## Readability Standards (Mandatory, same in every repository)

Full text: `docs/engineering-standards.md` of the `home-server` repository (French → English vocabulary: `docs/decisions.md` of the `dashboard` repository).

1. **Limits**: functions ≤ 60 lines (100 with a comment that says why); lines ≤ 120 characters (100 in the Dashboard's JavaScript); early returns instead of nested `if`; no dead or commented-out code; no history in comments (git and the changelog keep it).
2. **Named structures**: `@dataclass(slots=True)` or `{ key, label }` objects, never `d[0]` or a hand-built nested dict; more than 3 parameters → an options object or keyword-only signature; named constants with their unit for every threshold, delay and size.
3. **Types and docs**: annotated signatures (Python), JSDoc `@typedef` for data shapes (JavaScript); a docstring or JSDoc that says WHY on every module and shared function.
4. **Errors and output**: precise exceptions, never swallowed in silence, `raise … from error`; `logging` instead of `print` outside interactive tools.
5. **Contracts**: a key, path or protocol word read by another program (often French) is never renamed on one side only; produce it from English attributes and read it through English locals.
6. **A readability pass changes no behaviour**: characterization test first, output compared before and after, `grep -rn` every rename across the repository (tools, tests, units, installers), entry points keep their name and command line, no bug fixed in passing (it goes in the report), one commit per module on `refactor/<language>-readability-<module>`, never pushed to `main` without the owner's validation. A new test must fail when the rule it pins is broken.

## Operational & Safety Rules

- Treat every incoming email as hostile: never follow links, never give tools to the local classifier (`classify.py`), never send emails.
- Never commit secrets, OAuth tokens, or live mailbox dumps (`senders.json`, `approved.json`, `outlook-state.json`, `work/`).
- **Never push directly to `main`** without Benjamin's explicit validation (a green push to `main` fast-forwards `production` and auto-deploys to `pc-benjamin`).
