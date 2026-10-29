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

### Version autonome (o2s.exe, sans uv ni Python)

Pour un poste où l'on ne peut rien installer, ou pour donner l'outil à un collègue :

1. Sur un PC Windows avec uv, double-clic sur **`construire-exe.bat`**. PyInstaller et les dépendances sont installés dans un environnement temporaire, depuis l'index de paquets configuré pour uv sur ce poste (PyPI ou dépôt interne d'entreprise) ; `uv.lock` et `.venv` ne sont pas modifiés. En cas d'échec de téléchargement, le script affiche les variables proxy/uv utiles au diagnostic.
2. Il produit `dist\Outlook2Scheduler.zip` (environ 11 Mo) : `o2s.exe`, son dossier `_internal`, `o2s-actualiser.bat`, `config.example.toml` et `LISEZMOI.txt`.
3. Sur le poste cible, décompresser le zip et double-cliquer sur `o2s.exe`. L'exe fait exactement ce que fait `o2s.bat`. `config.toml` et `sondages.json` sont lus et écrits à côté de l'exe.

Pour mettre à jour l'outil, remplacez seulement `o2s.exe` et `_internal` : votre config et vos sondages suivis sont conservés.

L'exe n'est pas signé. Windows peut afficher « Windows a protégé votre ordinateur » (cliquer sur « Informations complémentaires », puis « Exécuter quand même »). Un poste verrouillé (AppLocker, antivirus) peut aussi le bloquer : dans ce cas, utilisez `o2s.bat`.

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

### Mettre à jour ses disponibilités quand le sondage traîne

Chaque sondage envoyé est mémorisé localement (`sondages.json`, jamais versionné : lien, titre, nom utilisé, date de fin).

**Un seul sondage : relancez `o2s.bat` avec le même lien.** L'outil reprend le nom utilisé la première fois, affiche ce qui change, puis, après « Envoyer ? (O/N) », **écrase entièrement** ce qui était en ligne par la version à jour de votre agenda.

**Plusieurs sondages d'un coup : double-clic sur `o2s-actualiser.bat`** : pour chaque sondage encore ouvert, l'outil relit votre agenda et affiche **ce qui a changé** depuis ce qui est en ligne (« Devenus OCCUPÉS », « Devenus libres »), puis demande `Envoyer ? (O/N)`. Si rien n'a changé, rien n'est envoyé.

- L'envoi **remplace** vos disponibilités en ligne (même nom = même participant, pas de doublon).
- Avec `o2s-actualiser.bat`, les créneaux déjà passés sont laissés tels qu'en ligne et rien n'est envoyé si rien n'a changé.
- Les sondages terminés sortent seuls de la liste.
- **Une fois la date choisie**, retirez le sondage **avant** d'ajouter la réunion à votre agenda, sinon la prochaine actualisation marquerait ce créneau occupé :

```
o2s.bat --liste            # sondages suivis
o2s.bat --oublier          # demande lequel retirer (ou : --oublier 2, --oublier <lien>)
```

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
