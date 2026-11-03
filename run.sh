#!/bin/zsh
# Daily pass over Outlook.com, started by launchd: runs daily-pass.mjs, with no agent and no browser window. `run.sh trial` is a dry run: it deletes, moves and recovers nothing.
# The installed launchd job runs this file by path: renaming it means updating the job. The notification is sent from here, and only when there is something to do.
MODE="${1:-auto}"
HERE="${0:A:h}"
LOG="$HOME/Library/Logs/outlook-spam-sorter.log"
print "=== $(date '+%Y-%m-%d %H:%M') run $MODE ===" >> "$LOG"
BEFORE=$(/opt/homebrew/bin/python3 -I -c 'import json,sys; print(json.load(open(sys.argv[1])).get("a_verifier",0))' "$HERE/outlook-state.json" 2>/dev/null || echo 0)
OUTPUT=$(mktemp)
/opt/homebrew/bin/node "$HERE/daily-pass.mjs" ${MODE:#auto} < /dev/null > "$OUTPUT" 2>&1
CODE=$?
cat "$OUTPUT" >> "$LOG"
SUMMARY=$(grep '^SUMMARY: ' "$OUTPUT" | tail -1); SUMMARY="${SUMMARY#SUMMARY: }"
if [[ $CODE -ne 0 || -z "$SUMMARY" ]]; then
  # Notification text: French on purpose.
  [[ -z "$SUMMARY" ]] && SUMMARY="Le passage a échoué, voir ~/Library/Logs/outlook-spam-sorter.log"
  [[ "$MODE" == "auto" ]] && "$HERE/record.py" --failure "$SUMMARY" > /dev/null
  ALERT=1
else
  # only when the number of mails to review goes up: otherwise the Dashboard tile is enough
  ALERT=$(/opt/homebrew/bin/python3 -I -c 'import json,sys; e=json.load(open(sys.argv[1])); print(1 if e.get("echec") or e.get("a_verifier",0)>int(sys.argv[2]) else 0)' "$HERE/outlook-state.json" "$BEFORE" 2>/dev/null || echo 1)
  [[ "$MODE" == "trial" ]] && ALERT=1
fi
[[ "$ALERT" == "1" ]] && /usr/bin/osascript -e 'on run argv' -e 'display notification (item 1 of argv) with title "Tri mails Outlook"' -e 'end run' "$SUMMARY"
rm -f "$OUTPUT"
find "$HERE/work" -mindepth 1 -delete
