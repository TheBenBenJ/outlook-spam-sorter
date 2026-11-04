# outlook-spam-sorter

A daily, unattended pass over an Outlook.com (Hotmail) mailbox that forwards to Gmail:

- it cleans the **junk folder**: known senders go back to the inbox, sure spam is deleted, the rest waits for a human;
- it recovers the mails **stuck in the inbox** (Outlook did not forward them) by depositing their original `.eml` into Gmail, then archives them;
- it adds the recovered senders to Outlook's **safe senders list**;
- it keeps an **activity journal** of what it did, mail by mail, for a dashboard (IMAP pass, see "The IMAP pass").

No AI drives the browser and no window opens: a script runs a headless Chrome. Every email is treated as hostile. A second version of the same pass, `imap-pass.py`, does it over IMAP with OAuth instead of a browser, so it can run on a server (see "The IMAP pass"). A Claude Code skill (`skill/`) covers what still needs a person: doubtful mails, unsubscribing, signing in again, repairing the script when Outlook's UI changes.

Formerly `tri-spam-outlook`: the project was renamed on 3 November 2026, along with its script and data file names (see "Design decisions").

## How it works

```
launchd (daily) ─► run.sh ─► daily-pass.mjs  (headless Chrome, persistent profile signed in to outlook.live.com)
                    │            │
                    │            ├─ Junk folder ── list snapshot ─► analyze.py --lines ─► classify.py (local model, no tools)
                    │            │     legitimate sender, or known in Gmail and judged legitimate ─► "Courrier légitime" (back to the inbox)
                    │            │     known spam domain, or unknown sender judged spam with confidence >= 0.8 ─► "Supprimer" (Deleted Items, 30 days)
                    │            │     anything else ─► left in place, counted as "to review"
                    │            ├─ Inbox ── list snapshot ─► analyze.py inbox ─► classify.py
                    │            │     trusted sender judged legitimate ─► download .eml ─► reinjecter (deposit into Gmail) ─► "Archiver"
                    │            │     already in Gmail, noise, old unknown ─► "Archiver"      recent unknown ─► left, "to review"
                    │            ├─ Safe senders list ── add the recovered senders
                    │            └─ record.py ─► senders.json, approved.json, outlook-state.json
                    └─ macOS notification, only if the run failed or the number of mails to review went up
```

| File | Role |
|---|---|
| `run.sh` | Entry point for launchd. Runs the pass, appends its report to `~/Library/Logs/outlook-spam-sorter.log`, sends the notification, empties `work/`. `run.sh trial` is a dry run. |
| `daily-pass.mjs` | The pass itself (Playwright, headless Chrome). Opens no link, sends nothing. |
| `imap-pass.py` | The same pass over IMAP (XOAUTH2), with no browser. `imap-pass.py trial` is a dry run. |
| `outlook_imap.py` | OAuth grant (`authorize <client-id>`, device code), token refresh, IMAP connection and folders (`check`). |
| `tests.py` | Offline checks of the IMAP pass and of `classify.py`: fake IMAP server, fake local model, fake deposit script, and a `claude` that fails the checks if anything calls it. |
| `analyze.py` | Reads a list snapshot and the settings, writes nothing: which class each row belongs to, what to recover, archive or approve. |
| `classify.py` | The only AI call, with no tools: a verdict from a closed list for each row, from the local model (`$OUTLOOK_SORTER_LOCAL_URL`, OpenAI API, default `127.0.0.1:11234`: the home server's Qwen 3.5 9B), in batches of 40. A row it does not judge gets no verdict and waits for the next run: no other model answers instead. `off`: no model at all, every row waits for a person. Each verdict names the model that gave it; one line on stderr says how many rows were judged and why not all. |
| `record.py` | Writes the outcome of a run into the three JSON files below, and prints the senders and spam domains it learned. |
| `activity.py` | Writer of the activity journal (see "The IMAP pass"): appends one JSON event per line, rotates the file above 5 MB. |
| `skill/` | Claude Code skill for interactive sessions (in French, like its trigger phrases). |
| `senders.json` | Settings and memory (git-ignored). Start from `senders.example.json`. |
| `approved.json` | Copy of Outlook's safe senders list (git-ignored). |
| `outlook-state.json` | State of the last run, for a dashboard or another service (git-ignored). |
| `activity.jsonl` | Activity journal of the IMAP pass, for a dashboard (git-ignored), with `activity-waiting.json`, the mails it already reported as waiting. |
| `.github/workflows/ci.yml` | Syntax checks on GitHub (Python, Node, zsh, JSON) for every push; no secret, no deployment. |

## The IMAP pass

`imap-pass.py` runs the same steps with the same helpers (`analyze.py`, `classify.py`, `record.py`) and the same files, over IMAP:

- **Authorization**: an app registration in Microsoft Entra (an Azure directory is required since June 2024, even for a personal account), with "Accounts in any organizational directory and personal Microsoft accounts" and public client flows allowed, so there is no client secret. `outlook_imap.py authorize <client-id>` prints a code to enter at microsoft.com/link from any device; the refresh token goes to `$OUTLOOK_SORTER_TOKEN` (default `~/.config/outlook-spam-sorter/token.json`, mode 600). Microsoft renews it at each use; unused for 90 days, it expires and the pass fails with "Session Outlook expirée".
- **Junk folder**: rows come from the headers (`From`, `Subject`, received date). Restoring is a move to the inbox, deleting a move to Deleted Items; the script never sets a deletion flag nor expunges.
- **Inbox**: the original `.eml` is fetched by IMAP (`BODY.PEEK[]`, the mail stays unread), handed to `reinjecter` (a path, or a command as a list), then the mail moves to the archive folder.
- **Not done**: adding recovered senders to Outlook's safe senders list. Neither IMAP nor Microsoft Graph exposes that list for a personal account, and Graph's `reportMessage` with `notJunk` (beta) moves the mail back to the inbox without adding its sender (tried on 3 November 2026). Only the web UI's "Courrier légitime" does. The senders already approved stay approved, and a known sender that lands in junk is still moved back by the next run.
- **State folder**: `$OUTLOOK_SORTER_DATA`, if set, holds the pass's state apart from its code: `senders.json`, `approved.json`, `outlook-state.json`, the activity journal and `work/` (on the home server, `/var/lib/routines/data/outlook-sorting`, next to the other sortings' data). By default they stay next to the scripts, as the browser pass expects.
- **Notification**: `$OUTLOOK_SORTER_NOTIFY`, if set, is called as `<command> "Tri mails Outlook" "<text>"` when the run fails or the number of mails to review goes up. The text carries counts or a fixed reason, never an address or a subject.
- **Activity journal**: `activity.jsonl` (`$OUTLOOK_SORTER_ACTIVITY`, by default next to `senders.json`), the home server's convention for the Dashboard's routines (home-server `docs/architecture.md`): one JSON event per line, `{"time", "action", "object", "from", "to", "detail", "engine", "model", "confidence", "run"}`, empty fields left out, texts in French because the owner reads them. One event per mail acted on: `remis en boîte`, `spam supprimé`, `à vérifier`, `récupéré dans Gmail`, `archivé`; per failure, `échec` (of the pass; of the local model, with how many mails it left without a verdict and why; or of one mail's deposit); per sender or domain learned, `ajouté aux légitimes`, `ajouté aux domaines spam`; and `passage terminé` with the run's counts. The object is the mail's subject, the detail the sender and why (`analyze.py`'s `motif` and the model's verdict); `engine`, `model` and `confidence` say which model decided, when one did. A mail left waiting is reported once per verdict, not at every hourly run (`activity-waiting.json`). The file holds subjects and addresses: mode 600, never in a notification. Above 5 MB it is rewritten with the last 90 days, 20,000 lines at most. Writing never fails a run, and a dry run writes nothing.

Checks: `python3 -I tests.py` (offline, writes only to a temporary folder).

## Requirements

- macOS (launchd, `osascript` notifications). The scripts call `/opt/homebrew/bin/node` and `/opt/homebrew/bin/python3`.
- Google Chrome, Node.js 20+, Python 3.
- A local model behind an OpenAI-compatible API (`$OUTLOOK_SORTER_LOCAL_URL`, default `http://127.0.0.1:11234`; for example llama.cpp's `llama-server` with Qwen 3.5 9B), for `classify.py`. Without one, every row that needs a verdict waits for a person. [Claude Code](https://docs.claude.com/en/docs/claude-code) only for the interactive skill.
- An Outlook.com account with the **French** web interface: the script clicks buttons by their French labels.
- A deposit script of your own (see `reinjecter` below) that puts an `.eml` file into Gmail.

## Install

```sh
git clone https://github.com/TheBenBenJ/outlook-spam-sorter ~/.local/share/outlook-spam-sorter && cd ~/.local/share/outlook-spam-sorter
npm ci
cp senders.example.json senders.json   # then edit: your addresses, browser profile, deposit script
mkdir -m 700 work
ln -s "$PWD/skill" ~/.claude/skills/outlook-spam-sorter
```

Sign in once to outlook.live.com in the browser profile named by `profil` (for example through the skill and the Playwright browser), then start with a dry run: `./run.sh trial` deletes, moves and recovers nothing.

Schedule `run.sh` with a launchd agent, for example `~/Library/LaunchAgents/com.example.outlook-spam-sorter.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.example.outlook-spam-sorter</string>
  <key>ProgramArguments</key><array><string>/bin/zsh</string><string>/Users/you/.local/share/outlook-spam-sorter/run.sh</string></array>
  <key>EnvironmentVariables</key><dict><key>PATH</key><string>/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin</string></dict>
  <key>StartCalendarInterval</key><dict><key>Hour</key><integer>7</integer><key>Minute</key><integer>30</integer></dict>
  <key>StandardOutPath</key><string>/dev/null</string>
  <key>StandardErrorPath</key><string>/Users/you/Library/Logs/outlook-spam-sorter.err.log</string>
</dict></plist>
```

`launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.example.outlook-spam-sorter.plist` loads it. Changing the scripts needs no reload: launchd only runs `run.sh`.

## Configuration

The settings and the JSON vocabulary are in French (see "Design decisions"). `senders.json`:

| Key | Meaning |
|---|---|
| `legitimes` | Addresses or domains always treated as legitimate (`address: note`). `record.py` adds addresses already received in Gmail. |
| `spam_confirme_domaines` | Domains whose mails are deleted without asking. `record.py` adds a domain after 3 confident spam verdicts in one run, never a consumer mail provider or a legitimate domain. |
| `desabonnes` | Newsletters you unsubscribed from (kept by the skill). |
| `moi` | Your own addresses: mails from them are archived, never recovered. |
| `gmail_connus` | Path to a JSON file kept up to date by your Gmail side (see below). Optional, but without it few senders count as known. |
| `libelles_bruit` | Gmail label prefixes for noise (advertising...): such mails are archived without recovery. |
| `libelles_non_fiables` | Gmail label prefixes that do not prove a sender is trusted. |
| `recuperer_depuis` | Older mails are archived, not recovered (`YYYY-MM-DD`). |
| `profil` | Chrome user data directory that holds the Outlook session. |
| `reinjecter` | Executable that deposits an `.eml` into Gmail (contract below). |
| `historique` | Last 120 runs, appended by `record.py`. |

`gmail_connus` points to a JSON file with `adresses` and `domaines` (senders already received in Gmail, with their mail count), `fournisseurs` (consumer mail providers, whose domain proves nothing), `libelles` and `exacts` (Gmail label of each domain and address), and `recents` (`[address, subject, YYYY-MM-DD]`, to recognise a mail Gmail already has).

**Reinjection contract.** `daily-pass.mjs` runs `reinjecter <file.eml>` and reads the last line of its output, a JSON object: `statut` is `depose` (deposited), `deja` (Gmail already has it), `refuse` (refused: the mail stays and counts as "to review") or `echec` (failed: retried on the next run); `de`, `objet` and `motif` are logged. The depositor must only accept a sender that is already known and whose mail is authenticated (for example: DMARC did not fail, and SPF or DKIM passed): a deposit bypasses Gmail's spam filter.

**State file.** `outlook-state.json` holds `maj` (time of the last run), `a_verifier` (mails to review), `recuperes`, `archives`, `restants` (left for the next run) and `echec` (failure reason, or `null`).

## Operations

- Dry run: `./run.sh trial`. Manual run: `launchctl kickstart -k gui/$(id -u)/<label>`. Only one browser at a time may use the profile.
- Report of each run: `~/Library/Logs/outlook-spam-sorter.log`. Rows waiting for a decision appear as `to review:` and `refused:` lines; the last line is the `SUMMARY:` used for the notification. launchd's own errors go to the `StandardErrorPath` file.
- Mails to review: open a Claude Code session and ask to sort the junk mail; the skill handles it with your approval.

Troubleshooting (the reasons below appear in the notification and in `outlook-state.json`):

| Message | What to do |
|---|---|
| `Session Outlook expirée, lancer /outlook-spam-sorter` | Sign in again in the profile (skill session), close that browser, run `run.sh`. |
| `Navigateur déjà ouvert sur le profil Outlook : passage sauté` | Another browser uses the profile (an interactive session): close it, run again. |
| `bouton « … » introuvable`, or many `action failed:` lines | Outlook's UI changed: run `./run.sh trial` and adapt the labels in `daily-pass.mjs`. |
| `Outlook propose une suppression définitive : annulée, passage arrêté` | A permanent-deletion dialog appeared; the script cancelled it. Find out why before the next run. |
| `Le passage Outlook dépasse 30 minutes : arrêté` | The watchdog stopped a slow run; the next run picks up what is left. |
| `local model qwen3.5-9b: 0 of N rows judged, the others wait for the next run: …` in the log | The local model did not answer (stopped, busy for more than 5 minutes, or `OUTLOOK_SORTER_LOCAL_URL=off`), or gave an unreadable answer: those rows wait for the next run, each reported once as « à vérifier », and the pass logs an « échec » with the reason; listed senders and spam domains are still handled. `model verdicts unavailable:` means `classify.py` itself failed. |
| `deposit failed:` in the log | The deposit script failed: the mail stays in the inbox and is retried next time. |

## Security notes

- Mail content is hostile. The script's actions never depend on what a mail says: it opens no link, sends nothing, clicks only buttons whose label matches exactly, never "Vider le dossier" (empty folder), and cancels any permanent-deletion dialog.
- The model sees only the sender address and the list row, has no tools, and answers from a closed list through a strict JSON schema (one answer per row, temperature 0, no thinking); `classify.py` keeps only the rows it asked, the three verdicts and a confidence between 0 and 1. A crafted mail can at worst leave a mail in place, or get the mail of an unknown sender deleted, which stays 30 days in Deleted Items. A known sender is never deleted.
- Only an already known sender is recovered into Gmail or added to the safe senders list: the inbox also holds phishing that got through Outlook's filters.
- The Chrome profile holds a signed-in session that can delete mail: keep it private, and never sign another account (Google, for example) into it.
- Downloaded `.eml` files go to `work/eml` (mode 700), which `run.sh` empties after every run. The JSON files hold addresses and are git-ignored, like the activity journal, which also holds subjects.
- The Python helpers run with `python3 -I` (isolated mode): no `PYTHON*` variables, no user site-packages, no imports from the current directory.

## Design decisions

- **A script drives the browser, not an agent.** The first version (October 2026) was a weekly Claude agent in `dontAsk` mode, with a closed tool list and Playwright MCP. Those limits reduced the risk of prompt injection without removing it: an agent reading hostile mail still held a signed-in browser that could delete messages. The pass is now a plain Playwright script whose actions depend only on sender addresses and fixed rules.
- **The local model decides, without Claude** (3-4 November 2026). On the home server, Qwen 3.5 9B answers the same closed list for free, and senders and subjects stay there. It read "confiance" as how legitimate a mail is, so one line of its prompt says what it means; it is also more categorical than Haiku (a compromised relay's mail that Haiku called doubtful, it calls spam, which sends it to Deleted Items, kept 30 days, instead of leaving it to review). It was first asked before Haiku, then compared with Haiku on real mail while Haiku decided (shadow mode, 3 November 2026: about 24 distinct mails, 3 disagreements, none where it would have deleted a mail Haiku judged legitimate). Since 4 November 2026 it is the only model: no fallback to Haiku, so the pass no longer depends on a subscription's usage window (on 3 November, four hourly passes went without a verdict while the usage limit was reached). When it does not answer, nothing moves and the rows wait for the next run. A model can only answer the rows it was asked, from the list.
- **One AI call, without tools, on a closed list.** Classifying a row is the only fuzzy step, so it is the only place a model is used. Deleting needs a spam verdict with confidence >= 0.8, and a missing verdict moves nothing.
- **Deposit the original instead of forwarding.** A forward arrives as "TR:", from the Hotmail address and dated today. Depositing the `.eml` over IMAP keeps the real sender, subject and date, and sends nothing. Because a deposit skips Gmail's spam filter, the depositor accepts only known, authenticated senders.
- **Why sweep the inbox at all.** Outlook forwards to Gmail without keeping a copy, except for a few percent of mails that stay in the inbox, including mails from safe senders. Those mails carry a spam confidence level of 5 (`X-MS-Exchange-Organization-SCL`): the safe senders list makes them land in the inbox, not leave it. What is left in the inbox therefore never reached Gmail.
- **"Courrier légitime" instead of moving mails out of junk.** "Archiver" is disabled in the junk folder. "Courrier légitime" (Not junk) sends the mail back to the inbox, where the same run recovers it, and its dialog (ticked by default) adds the sender to the safe senders list.
- **Clicks through JavaScript, on exact labels.** An Outlook overlay sometimes intercepts simulated clicks. Outlook's dialogs are fixed-position, so visibility is tested with `getClientRects()`, not `offsetParent`.
- **Bounded runs.** Rows are ticked in batches of 25 and the batch is dropped if anything else is selected; a run recovers at most 100 mails and archives at most 300; a 30-minute watchdog stops it. The rest waits for the next day.
- **Notify only when something changed.** A notification appears when a run fails or the number of mails to review goes up; otherwise the state file (shown on a dashboard) is enough.
- **Microsoft Graph was considered and set aside.** It would replace the browser with an API (`$value` for the original `.eml`, `move`, `Mail.ReadWrite`), but a personal account needs an app registration in an Azure directory, `mailboxSettings` does not expose the safe senders list, and `markAsNotJunk` is deprecated (its successor, `reportMessage`, is beta only).
- **IMAP with OAuth, November 2026.** Tried on the real mailbox on 3 November 2026: with an app registration in an Azure directory, a personal account grants `IMAP.AccessAsUser.All` by device code, and IMAP reads the folders, fetches the original `.eml` and moves mails (`MOVE`). It drops the browser, the Chrome profile, Node and the UI selectors, and runs on a server. The price is an Azure directory, and the safe senders list, which nothing but the web UI can change (Graph's `reportMessage notJunk` does not add the sender either: tried the same day). `imap-pass.py` is that pass; the browser pass stays until the switch.
- **Real Chrome with a persistent profile.** `playwright-core` (pinned, no bundled browser) drives the installed Chrome headless, so the Outlook session survives between runs and the interactive skill can reuse the same profile (one browser at a time).
- **What stays in French, and why.** Code, comments, logs and docs are in English, and so are the names of the scripts and data files. French stays where it is an interface: Outlook's French button labels (selectors), the notification and failure texts (read by a person; a dashboard matches "Session Outlook"), the model's prompt and its verdict values (rewording would change the verdicts), the JSON keys and class names shared with other components (`legitimes`, `a_verifier`, `recuperer`...), the reinjection statuses, and the skill, whose trigger phrases must match French requests.
- **Renamed on 3 November 2026.** The project was `tri-spam-outlook`; its folder, launchd label, log and skill now use `outlook-spam-sorter`. `lancer.sh` became `run.sh` and its dry-run argument `essai` became `trial`; any other argument is refused, so the old word cannot start a real pass. The data files became `senders.json` (formerly `expediteurs.json`), `approved.json` (`approuves.json`) and `outlook-state.json` (`etat-outlook.json`), with the same keys and values: the programs that read them and the launchd job switched at the same time. `.gitignore` still lists the old names, so a copy from before is never committed.

## License

MIT, see [LICENSE](LICENSE).
