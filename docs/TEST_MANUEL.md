# Test manuel sur le poste (Outlook réel + When2meet bidon)

Les tests automatiques simulent Outlook et When2meet. Cette procédure valide les deux parties non testables automatiquement : **la lecture COM** et **l'envoi réel**. Durée : ~20 min.

Notez pour chaque étape : OK / KO + le message affiché.

## 0. Préparation

1. Cloner le dépôt, copier `config.example.toml` en `config.toml`, mettre `nom = "Test o2s"`.
2. Ouvrir une invite de commandes dans le dossier et lancer `uv run --system-certs pytest` (ou `--native-tls` sur un uv ancien).
   Attendu : tous les tests passent. Valide l'installation à travers le proxy.
3. Lancer `o2s.bat --test-reseau`. Attendu : `mode windows : OK`. Le mode python peut échouer derrière un proxy à configuration automatique : c'est normal.
4. Dans Outlook, choisir un jour ouvré **J** de la semaine qui vient, **sans autre rendez-vous entre 9:00 et 17:00** (sinon, adapter les résultats attendus), et créer deux rendez-vous **sans participant** :
   - « Test A » de 10:00 à 10:30, statut **Occupé** ;
   - « Test B » de 14:00 à 15:00, statut **Provisoire**.

## 1. Lecture de l'agenda (COM)

```
o2s.bat --agenda AAAA-MM-JJ        (date de J)
```

Attendu (une ligne par heure, un caractère par quart d'heure) :
```
  10:00  XX..
  14:00  pppp
```
- Alerte « Un programme tente d'accéder… » : cliquer **Autoriser**, puis noter si elle revient à chaque lancement.
- Décalage d'une ou deux heures, ou mauvais jour : **KO à me remonter** (problème de fuseau dans l'appel COM).
- Outlook fermé : relancer après avoir fermé Outlook → message « Outlook ne semble pas ouvert ».

## 2. Sondage bidon

Sur https://www.when2meet.com, créer « Test o2s » :
- dates : **J**, le **samedi** suivant, et un jour à **J + 3 semaines** (dates non consécutives) ;
- heures : 9:00 à 17:00, fuseau Europe/Paris.

Copier le lien (Ctrl+C dans la barre d'adresse).

## 3. Aperçu puis refus

Double-cliquer sur `o2s.bat`.

Attendu :
- « Lien lu dans le presse-papiers » ;
- jour J : `09:00-10:00, 10:30-14:00, 15:00-17:00` ;
- samedi : `-- aucune --` ;
- J + 3 semaines : selon votre agenda réel ;
- `Envoyer ? (O/N)` → répondre **N** → « Envoi annulé ».

Rafraîchir le sondage dans le navigateur : aucun participant « Test o2s ».

## 4. Envoi réel

Relancer `o2s.bat`, répondre **O**.

Attendu : « OK : N créneau(x) libres enregistrés et vérifiés sur la page ».
Dans le navigateur : « Test o2s » apparaît, sa grille correspond à l'aperçu.

Si le message est « ATTENTION : l'envoi a été accepté mais la page ne reflète pas… » : **KO à me remonter** (When2meet n'accepte pas l'envoi direct).

## 5. Options et remplacement

```
o2s.bat --provisoire-libre
```
Attendu : jour J `09:00-10:00, 10:30-17:00`. Envoyer (O), rafraîchir : la grille de « Test o2s » est **remplacée** (14:00-15:00 maintenant libre).

```
o2s.bat --marge 15 --heures 09:30-16:00
```
Attendu : jour J `09:30-09:45, 10:45-13:45, 15:15-16:00` (aperçu seulement, répondre N).

## 6. Messages d'erreur

| Action | Attendu |
| --- | --- |
| Dans le navigateur, se connecter au sondage comme « Test mdp » avec le mot de passe `abc`, puis `o2s.bat --nom "Test mdp" --mot-de-passe faux` et O | « mot de passe refusé » |
| Copier un texte quelconque puis double-clic sur `o2s.bat` | le lien est demandé |
| Créer un sondage « Days of the week » et lancer dessus | refus « jours de la semaine » |
| Lien d'un sondage à + 4 mois | avertissement « hors de la grille Outlook » (selon l'horizon publié) |
| Débrancher le réseau / Wi-Fi | « Impossible de joindre When2meet » |

## 7. Changement d'heure

Prochain passage : dimanche 28/03/2027.

1. Créer « Test C » le **lundi 29/03/2027** de 10:00 à 11:00 (Occupé).
2. `o2s.bat --agenda 2027-03-29` → ligne `10:00  XXXX`.
3. Sondage bidon sur vendredi 26/03 et lundi 29/03/2027, 9:00-12:00 : l'aperçu doit montrer lundi `09:00-10:00, 11:00-12:00`.

## 8. Nettoyage

Supprimer les rendez-vous « Test A/B/C ». Les sondages When2meet ne se suppriment pas : ils expirent seuls.
