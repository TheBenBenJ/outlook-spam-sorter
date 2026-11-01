# tri-spam-outlook

A script that cleans an Outlook.com junk folder every day and moves the mails stuck in its inbox to Gmail, unattended, without a window, and without any AI driving the browser. Every email is treated as hostile. A Claude Code skill covers what still needs a human: doubtful mails, unsubscribing, signing in again.

Un script qui vide chaque jour le dossier indésirable d'Outlook.com et récupère dans Gmail les mails restés coincés dans sa boîte de réception, sans personne devant l'écran, sans fenêtre, et sans IA au volant du navigateur. Chaque mail est traité comme hostile. Un skill Claude Code couvre ce qui demande encore quelqu'un : les mails douteux, les désinscriptions, la reconnexion.

## Garanties

- `passage.mjs` est un script : ce qu'il fait ne dépend pas de ce que disent les mails. Il n'ouvre aucun lien et n'envoie rien.
- `classer.py` est le seul appel à une IA : Haiku, sans aucun outil, ne voit que l'adresse de l'expéditeur et la ligne de la liste, et répond dans une liste fermée. Un mail piégé peut au pire laisser un mail en place, ou faire supprimer le mail d'un expéditeur inconnu, qui reste 30 jours dans Éléments supprimés.
- Un expéditeur déjà connu n'est jamais supprimé. Seul un expéditeur déjà connu revient en boîte ou entre sur la liste blanche.
- La récupération passe par le téléchargement de l'original en .eml, déposé dans Gmail par un script à fournir (`reinjecter` dans les réglages), qui ne doit accepter qu'un expéditeur déjà connu et authentifié.
- Si Outlook propose une suppression définitive, le script annule et s'arrête.

`lancer.sh essai` ne supprime, ne déplace et ne récupère rien : commence par là.

## Installation

```sh
git clone https://github.com/TheBenBenJ/tri-spam-outlook ~/.local/share/tri-spam-outlook && cd ~/.local/share/tri-spam-outlook
npm ci
cp expediteurs.exemple.json expediteurs.json   # expéditeurs légitimes, profil du navigateur, script de dépôt
mkdir travail
ln -s "$PWD/skill" ~/.claude/skills/tri-spam-outlook
```

Connecte-toi une fois à outlook.live.com dans le profil indiqué par `profil`, puis planifie `lancer.sh` (launchd, cron). Il faut Google Chrome et Claude Code.

`gmail_connus` pointe vers un JSON tenu à jour par ta boîte de destination : `adresses` et `domaines` (expéditeurs déjà reçus, avec leur nombre de mails), `fournisseurs` (messageries grand public, dont le domaine ne prouve rien), `libelles` et `exacts` (libellé de chaque domaine et adresse), `recents` (`[adresse, objet, jour]`, pour reconnaître ce qui est déjà arrivé).

## Limites

Les sélecteurs suivent l'interface française d'Outlook et peuvent casser quand elle change : le passage s'arrête alors et notifie, sans rien faire de plus. Le navigateur garde une session connectée à ta boîte mail, et le script peut y supprimer des messages.
