# tri-spam-outlook

A Claude Code skill that cleans an Outlook.com junk folder every week, unattended, built on the assumption that every email is hostile: closed tool list, no free shell, no file writes, no links opened, and a trusted-senders list that an email can never modify.

Un skill Claude Code qui vide chaque semaine le dossier indésirable d'Outlook.com, sans personne devant l'écran, en partant du principe que chaque mail est hostile.

## Garanties du mode automatique

- `lancer.sh` lance Claude en `dontAsk`, avec une liste fermée d'outils. Il n'y a ni shell libre, ni écriture de fichier, ni `browser_run_code_unsafe`.
- `analyser.py` et `consigner.py` sont les seuls scripts autorisés. `consigner.py` ne touche jamais à `legitimes` : un mail piégé ne peut pas y faire entrer un expéditeur.
- Le classement se fait sur l'adresse réelle et le libellé de la liste, sans ouvrir un mail pour le lire ni suivre de lien.

Le mode interactif montre le classement et attend ton accord. `lancer.sh essai` ne supprime et ne déplace rien : commence par là.

## Installation

```sh
git clone https://github.com/TheBenBenJ/tri-spam-outlook ~/.local/share/tri-spam-outlook && cd ~/.local/share/tri-spam-outlook
cp expediteurs.exemple.json expediteurs.json   # à remplir avec tes expéditeurs légitimes
cp mcp.exemple.json mcp.json                   # à adapter
mkdir travail
ln -s "$PWD/skill" ~/.claude/skills/tri-spam-outlook
```

Connecte-toi une fois à outlook.live.com dans le profil Playwright, puis planifie `lancer.sh` (launchd, cron). Les sélecteurs suivent l'interface d'Outlook et peuvent casser quand elle change.

Ces garanties réduisent le risque d'injection de prompt sans l'annuler. L'agent garde un navigateur connecté à ta boîte mail, et il peut y supprimer des messages.
