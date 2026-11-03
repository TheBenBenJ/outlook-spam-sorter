#!/bin/zsh
# Passage quotidien sur Outlook.com (launchd) : daily-pass.mjs, sans agent ni fenêtre. `lancer.sh essai` ne supprime, ne déplace et ne récupère rien.
# La notification part d'ici, seulement s'il y a quelque chose à faire.
MODE="${1:-auto}"
ICI="${0:A:h}"
LOG="$HOME/Library/Logs/tri-spam-outlook.log"
print "=== $(date '+%Y-%m-%d %H:%M') passage $MODE ===" >> "$LOG"
AVANT=$(/opt/homebrew/bin/python3 -I -c 'import json,sys; print(json.load(open(sys.argv[1])).get("a_verifier",0))' "$ICI/etat-outlook.json" 2>/dev/null || echo 0)
SORTIE=$(mktemp)
/opt/homebrew/bin/node "$ICI/daily-pass.mjs" ${MODE:#auto} < /dev/null > "$SORTIE" 2>&1
CODE=$?
cat "$SORTIE" >> "$LOG"
RESUME=$(grep '^RÉSUMÉ : ' "$SORTIE" | tail -1); RESUME="${RESUME#RÉSUMÉ : }"
if [[ $CODE -ne 0 || -z "$RESUME" ]]; then
  [[ -z "$RESUME" ]] && RESUME="Le passage a échoué, voir ~/Library/Logs/tri-spam-outlook.log"
  [[ "$MODE" == "auto" ]] && "$ICI/record.py" --echec "$RESUME" > /dev/null
  ALERTE=1
else
  # seulement si le nombre de mails à vérifier augmente : sinon la tuile du Dashboard suffit
  ALERTE=$(/opt/homebrew/bin/python3 -I -c 'import json,sys; e=json.load(open(sys.argv[1])); print(1 if e.get("echec") or e.get("a_verifier",0)>int(sys.argv[2]) else 0)' "$ICI/etat-outlook.json" "$AVANT" 2>/dev/null || echo 1)
  [[ "$MODE" == "essai" ]] && ALERTE=1
fi
[[ "$ALERTE" == "1" ]] && /usr/bin/osascript -e 'on run argv' -e 'display notification (item 1 of argv) with title "Tri mails Outlook"' -e 'end run' "$RESUME"
rm -f "$SORTIE"
find "$ICI/work" -mindepth 1 -delete
