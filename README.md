# Outlook2Scheduler

Remplit automatiquement vos sondages de planification à partir de votre agenda **Outlook classique** (Windows). Pris en charge : **When2meet** (Framadate et Doodle prévus).

- Lecture de l'agenda Outlook par COM, sans Microsoft Graph ni extension de navigateur : seuls l'heure de début, l'heure de fin et le statut (libre, provisoire, occupé…) de chaque rendez-vous sont lus.
- **Confidentialité** : seule une grille libre/occupé (0/1) est envoyée. Aucun intitulé, participant ou contenu de réunion n'est lu.
- Aperçu des plages libres et confirmation **« Envoyer ? (O/N) »** avant tout envoi, puis vérification que When2meet a bien enregistré.

## Installation

1. Installer [uv](https://docs.astral.sh/uv/) (PowerShell) :
   ```
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```
2. Récupérer ce dossier (clone ou zip).
3. Copier `config.example.toml` en `config.toml` et y mettre votre nom et votre plage horaire.

Au premier lancement, uv installe Python et les dépendances (`pywin32`, `requests`, `truststore`) dans `.venv`.

### Proxy d'entreprise

- `o2s.bat` demande à uv d'utiliser les certificats Windows (`UV_SYSTEM_CERTS`, ou `UV_NATIVE_TLS` pour un uv ancien). En ligne de commande : `uv run --system-certs ...` (ou `--native-tls`).
- Pour joindre When2meet, l'outil passe par défaut par **la pile réseau de Windows** (celle d'Edge et d'Excel) : mêmes réglages proxy que le navigateur (y compris script de configuration automatique), authentification au proxy avec votre session Windows, certificats Windows.
- Vérifier l'accès sans rien envoyer : `o2s.bat --test-reseau`.
- Autre mode : `--reseau python` (ou `reseau = "python"` dans `config.toml`) utilise `requests` + `truststore` ; il ne gère que les proxys fixes (`HTTPS_PROXY`).
- Si uv ne peut pas **télécharger Python**, utilisez le Python déjà installé (3.11 ou plus) :
  ```
  uv python find
  uv venv --python "C:\chemin\vers\python.exe"
  ```
  ou décommentez la ligne `set "UV_PYTHON=..."` dans `o2s.bat`.

## Utilisation

**Double-clic sur `o2s.bat`** après avoir copié le lien du sondage :

1. le lien est lu dans le presse-papiers (sinon il est demandé) ;
2. le sondage et votre agenda sont lus, l'aperçu des plages libres s'affiche ;
3. `Envoyer ? (O/N)` : rien ne part sans un « O ».

Ligne de commande (même options via le `.bat`) :

```
uv run o2s <lien> [--nom "Prénom N."] [--heures 08:30-18:00] [--jours-exclus sam,dim]
               [--provisoire-libre] [--marge 15] [--mot-de-passe X] [--submit] [--oui]
uv run o2s --agenda 2026-11-02     # diagnostic : grille Outlook d'un jour, rien n'est envoyé
uv run o2s --test-reseau           # diagnostic : accès à When2meet, rien n'est envoyé
```

Sans `--submit`, l'outil affiche seulement l'aperçu.

### Règles de disponibilité

Un créneau de 15 min est « libre » s'il est dans la plage horaire, hors jours exclus, et qu'Outlook le donne libre sur toute sa durée (et sur la marge éventuelle). « Provisoire » est occupé par défaut, « absent » toujours occupé, « travaille ailleurs » libre par défaut (`ailleurs_libre`).

## Limites

- Outlook **classique** uniquement (pas le nouvel Outlook), et il doit être **ouvert**.
- Seul le calendrier principal est lu (pas les calendriers partagés ou secondaires).
- Ancienne méthode disponible : `--source freebusy` (grille libre/occupé d'Outlook). Elle déclenche sur certains postes l'alerte « Un programme tente d'accéder… » ; elle est limitée à l'horizon publié et signale le jour de changement d'heure.
- Les sondages « jours de la semaine » (sans dates) sont refusés.
- When2meet n'a pas d'API officielle : l'outil utilise les mêmes requêtes que la page. Si le site change, un message « format de page changé » s'affiche.
- **Envoyer remplace** vos disponibilités précédentes dans ce sondage.

## Développement

```
uv run pytest
```

Validation sur un vrai poste (Outlook + When2meet bidon) : [docs/TEST_MANUEL.md](docs/TEST_MANUEL.md).

Architecture : `sources/` (agenda), `cibles/` (un module par outil de sondage), `regles.py`, `apercu.py`, `cli.py`. Framadate et Doodle sont prévus comme nouvelles cibles.

Licence MIT.
