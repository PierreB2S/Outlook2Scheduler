# outlook2when2meet

Remplit automatiquement un sondage **When2meet** à partir de votre agenda **Outlook classique** (Windows).

- Lecture de l'agenda par COM (`Recipient.FreeBusy`), sans Microsoft Graph, sans extension de navigateur.
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

- `w2m.bat` demande à uv d'utiliser les certificats Windows (`UV_SYSTEM_CERTS`, ou `UV_NATIVE_TLS` pour un uv ancien). En ligne de commande : `uv run --system-certs ...` (ou `--native-tls`).
- Côté Python, `truststore` fait utiliser les certificats Windows pour joindre When2meet.
- Si uv ne peut pas **télécharger Python**, utilisez le Python déjà installé (3.11 ou plus) :
  ```
  uv python find
  uv venv --python "C:\chemin\vers\python.exe"
  ```
  ou décommentez la ligne `set "UV_PYTHON=..."` dans `w2m.bat`.

## Utilisation

**Double-clic sur `w2m.bat`** après avoir copié le lien du sondage :

1. le lien est lu dans le presse-papiers (sinon il est demandé) ;
2. le sondage et votre agenda sont lus, l'aperçu des plages libres s'affiche ;
3. `Envoyer ? (O/N)` : rien ne part sans un « O ».

Ligne de commande (même options via le `.bat`) :

```
uv run w2m <lien> [--nom "Prénom N."] [--heures 08:30-18:00] [--jours-exclus sam,dim]
               [--provisoire-libre] [--marge 15] [--mot-de-passe X] [--submit] [--oui]
uv run w2m --agenda 2026-11-02     # diagnostic : grille Outlook d'un jour, rien n'est envoyé
```

Sans `--submit`, l'outil affiche seulement l'aperçu.

### Règles de disponibilité

Un créneau de 15 min est « libre » s'il est dans la plage horaire, hors jours exclus, et qu'Outlook le donne libre sur toute sa durée (et sur la marge éventuelle). « Provisoire » est occupé par défaut, « absent » toujours occupé, « travaille ailleurs » libre par défaut (`ailleurs_libre`).

## Limites

- Outlook **classique** uniquement (pas le nouvel Outlook), et il doit être **ouvert**.
- Si Outlook affiche « Un programme tente d'accéder… », cliquez sur **Autoriser**. Si l'accès est bloqué sans alerte, c'est la politique de sécurité du poste : voir avec l'IT.
- La grille vient des informations libre/occupé d'Outlook : au-delà de l'horizon publié (souvent 2 mois), les créneaux sont « hors grille » et marqués occupés.
- Les sondages « jours de la semaine » (sans dates) sont refusés.
- Le **jour même** d'un changement d'heure est signalé dans l'aperçu : vérifiez-le à la main.
- When2meet n'a pas d'API officielle : l'outil utilise les mêmes requêtes que la page. Si le site change, un message « format de page changé » s'affiche.
- Proxy configuré par script PAC : non géré par `requests` ; définir `HTTPS_PROXY` si besoin.
- **Envoyer remplace** vos disponibilités précédentes dans ce sondage.

## Développement

```
uv run pytest
```

Architecture : `sources/` (agenda), `cibles/` (un module par outil de sondage), `regles.py`, `apercu.py`, `cli.py`. Framadate et Doodle sont prévus comme nouvelles cibles.

Licence MIT.
