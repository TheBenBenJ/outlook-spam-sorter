#!/bin/zsh
# Passage hebdomadaire du skill tri-spam-outlook (launchd) ; `lancer.sh essai` ne supprime et ne déplace rien.
# L'agent lit des mails hostiles sans personne derrière l'écran : dontAsk refuse tout ce qui n'est pas listé ici. Ni shell libre, ni écriture de fichier,
# ni browser_run_code_unsafe (équivaut à exécuter du code sur le Mac). Le dossier travail/ borne les fichiers écrits par Playwright. La notification part d'ici.
MODE="${1:-auto}"
ICI="$HOME/.local/share/tri-spam-outlook"
LOG="$HOME/Library/Logs/tri-spam-outlook.log"
cd "$ICI/travail" || exit 1
print "=== $(date '+%Y-%m-%d %H:%M') passage $MODE ===" >> "$LOG"
SORTIE=$(mktemp)
PW=mcp__playwright__browser
/opt/homebrew/bin/claude -p "/tri-spam-outlook $MODE" \
  --permission-mode dontAsk --strict-mcp-config --mcp-config "$ICI/mcp.json" \
  --tools Read Bash Skill \
  --allowedTools "Skill(tri-spam-outlook)" "Read(/$ICI/expediteurs.json)" \
    "Bash($ICI/analyser.py *)" "Bash($ICI/consigner.py *)" \
    ${PW}_navigate ${PW}_snapshot ${PW}_evaluate ${PW}_click ${PW}_type ${PW}_press_key ${PW}_wait_for ${PW}_find ${PW}_handle_dialog ${PW}_close \
  --disallowedTools ${PW}_run_code_unsafe ${PW}_file_upload \
  < /dev/null > "$SORTIE" 2>&1
CODE=$?
cat "$SORTIE" >> "$LOG"
RESUME=$(grep '^RÉSUMÉ : ' "$SORTIE" | tail -1); RESUME="${RESUME#RÉSUMÉ : }"
[[ $CODE -eq 0 && -n "$RESUME" ]] || RESUME="Le passage a échoué, voir ~/Library/Logs/tri-spam-outlook.log"
/usr/bin/osascript -e 'on run argv' -e 'display notification (item 1 of argv) with title "Tri spam Outlook"' -e 'end run' "$RESUME"
rm -f "$SORTIE"
find "$ICI/travail" -mindepth 1 -delete
