"""Ligne de commande : o2s <lien> [--nom] [--heures 08:30-18:00] [--provisoire-libre] [--submit]."""

from __future__ import annotations

import argparse
import sys
import time
import traceback
from dataclasses import dataclass, field, replace
from datetime import date, timedelta
from pathlib import Path
from typing import Callable

from outlook2scheduler import __version__
from outlook2scheduler import reseau
from outlook2scheduler.apercu import formater, formater_ecart
from outlook2scheduler.cibles.base import CibleSondage
from outlook2scheduler.config import Config, charger
from outlook2scheduler.erreurs import ErreurUtilisateur
from outlook2scheduler.registre import Registre
from outlook2scheduler.regles import Motif, Regles, evaluer, parse_heures, parse_jours
from outlook2scheduler.sources.base import PAS_SECONDES, SourceAgenda
from outlook2scheduler.temps import JOURS_COURTS, Fuseau, FuseauSysteme, minuit

CODES_LISIBLES = {"0": ".", "1": "p", "2": "X", "3": "A", "4": "a"}


def lire_presse_papiers() -> str:  # pragma: no cover - dépend de Windows
    try:
        import win32clipboard

        win32clipboard.OpenClipboard()
        try:
            return str(win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT) or "")
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return ""


@dataclass
class Dependances:
    """Tout ce qui touche au monde extérieur, remplaçable dans les tests."""

    fabrique_source: Callable[[Fuseau, str], SourceAgenda] | None = None
    fabrique_cibles: Callable[[str], list[CibleSondage]] | None = None
    fuseau: Fuseau = field(default_factory=FuseauSysteme)
    entree: Callable[[str], str] = input
    sortie: Callable[[str], None] = print
    presse_papiers: Callable[[], str] = lire_presse_papiers
    maintenant: Callable[[], float] = time.time


SOURCES = ("rendez-vous", "freebusy")


def _source_outlook(fuseau: Fuseau, source: str) -> SourceAgenda:  # pragma: no cover
    if source == "freebusy":
        from outlook2scheduler.sources.outlook import OutlookCOM, SourceOutlook

        return SourceOutlook(OutlookCOM(), fuseau)
    from outlook2scheduler.sources.outlook_rdv import OutlookRendezVousCOM, SourceOutlookRendezVous

    return SourceOutlookRendezVous(OutlookRendezVousCOM(), fuseau)


def _cibles_par_defaut(mode_reseau: str) -> list[CibleSondage]:  # pragma: no cover
    from outlook2scheduler.cibles.when2meet import When2meet

    return [When2meet(reseau.creer_session(mode_reseau))]


def construire_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="o2s",
        description="Remplit un sondage de planification (When2meet pour l'instant) à partir de votre agenda Outlook. "
        "Seule une grille libre/occupé est envoyée.",
    )
    p.add_argument("lien", nargs="?", help="lien du sondage (sinon : presse-papiers, puis question)")
    p.add_argument("--nom", help="votre nom dans le sondage")
    p.add_argument("--mot-de-passe", dest="mot_de_passe", help="mot de passe When2meet (facultatif)")
    p.add_argument("--heures", help="plage horaire acceptée, ex. 08:30-18:00")
    p.add_argument("--jours-exclus", dest="jours_exclus", help="ex. sam,dim (ou « aucun »)")
    p.add_argument(
        "--provisoire-libre",
        dest="provisoire_libre",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="compter les rendez-vous provisoires comme libres",
    )
    p.add_argument("--marge", type=int, dest="marge_minutes", help="marge en minutes avant/après chaque réunion")
    p.add_argument("--config", type=Path, help="fichier de configuration (défaut : config.toml)")
    p.add_argument("--submit", action="store_true", help="envoyer après aperçu et confirmation")
    p.add_argument("--oui", action="store_true", help="avec --submit : ne pas demander de confirmation")
    p.add_argument(
        "--actualiser",
        action="store_true",
        help="remet à jour tous les sondages déjà envoyés et encore ouverts (montre les changements)",
    )
    p.add_argument("--liste", action="store_true", help="liste les sondages suivis")
    p.add_argument(
        "--oublier",
        nargs="?",
        const="",
        metavar="N|LIEN",
        help="ne plus actualiser un sondage (ex. une fois la date choisie)",
    )
    p.add_argument(
        "--agenda",
        metavar="AAAA-MM-JJ",
        help="diagnostic : affiche la grille Outlook de ce jour (rien n'est envoyé)",
    )
    p.add_argument(
        "--source",
        choices=SOURCES,
        help="lecture d'Outlook : rendez-vous (défaut, évite l'alerte de sécurité) ou freebusy",
    )
    p.add_argument(
        "--reseau",
        choices=reseau.MODES,
        help="accès Internet : windows (comme Edge, gère le proxy d'entreprise ; défaut) ou python",
    )
    p.add_argument(
        "--test-reseau",
        action="store_true",
        help="diagnostic : teste l'accès à When2meet dans les deux modes (rien n'est envoyé)",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def regles_depuis(config: Config, args: argparse.Namespace) -> Regles:
    debut, fin = parse_heures(args.heures or config.heures)
    jours = parse_jours(args.jours_exclus if args.jours_exclus is not None else config.jours_exclus)
    provisoire = config.provisoire_libre if args.provisoire_libre is None else args.provisoire_libre
    marge = config.marge_minutes if args.marge_minutes is None else args.marge_minutes
    return Regles(
        heure_debut=debut,
        heure_fin=fin,
        jours_exclus=jours,
        provisoire_libre=provisoire,
        ailleurs_libre=config.ailleurs_libre,
        marge_minutes=marge,
    )


def _choisir_cible(lien: str, cibles: list[CibleSondage]) -> CibleSondage | None:
    return next((c for c in cibles if c.reconnait(lien)), None)


def _obtenir_lien(args: argparse.Namespace, cibles: list[CibleSondage], dep: Dependances) -> tuple[str, CibleSondage]:
    if args.lien:
        cible = _choisir_cible(args.lien, cibles)
        if cible is None:
            raise ErreurUtilisateur(f"Lien non reconnu : {args.lien!r} (seul When2meet est pris en charge).")
        return args.lien.strip(), cible

    presse = (dep.presse_papiers() or "").strip()
    cible = _choisir_cible(presse, cibles) if presse else None
    if cible is not None:
        dep.sortie(f"Lien lu dans le presse-papiers : {presse}")
        return presse, cible

    for _ in range(3):
        saisie = dep.entree("Collez le lien du sondage When2meet puis Entrée : ").strip().strip('"')
        cible = _choisir_cible(saisie, cibles)
        if cible is not None:
            return saisie, cible
        dep.sortie("Ce n'est pas un lien When2meet (ex. https://www.when2meet.com/?12345678-AbCdE).")
    raise ErreurUtilisateur("Aucun lien When2meet valide fourni.")


def _diagnostic_agenda(jour_texte: str, dep: Dependances, source: str) -> int:
    try:
        jour = date.fromisoformat(jour_texte)
    except ValueError as exc:
        raise ErreurUtilisateur(f"Date invalide : {jour_texte!r} (format AAAA-MM-JJ).") from exc
    assert dep.fabrique_source is not None
    grille = dep.fabrique_source(dep.fuseau, source).grille([jour])
    debut = dep.fuseau.vers_epoch(minuit(jour))
    fin = dep.fuseau.vers_epoch(minuit(jour + timedelta(days=1)))
    dep.sortie(f"Grille Outlook du {JOURS_COURTS[jour.weekday()]} {jour:%d/%m/%Y} "
               "(. libre, p provisoire, X occupé, A absent, a ailleurs, ? inconnu)")
    for heure_debut in range(debut, fin, 4 * PAS_SECONDES):
        quarts = range(heure_debut, min(heure_debut + 4 * PAS_SECONDES, fin), PAS_SECONDES)
        codes = "".join(
            "?" if (c := grille.code(q)) is None else CODES_LISIBLES.get(str(c), "?") for q in quarts
        )
        dep.sortie(f"  {dep.fuseau.vers_local(heure_debut):%H:%M}  {codes}")
    if grille.nb_rendez_vous is not None:
        dep.sortie(f"{grille.nb_rendez_vous} rendez-vous non libre(s) lu(s) ce jour-là (horaires et statut uniquement).")
    if jour in grille.jours_changement_heure:
        dep.sortie("Attention : jour de changement d'heure.")
    return 0


def _diagnostic_reseau(dep: Dependances) -> int:  # pragma: no cover - dépend du poste
    dep.sortie(f"Test d'accès à {reseau.URL_TEST} (simple lecture, rien n'est envoyé)")
    ok = False
    for mode in reseau.MODES:
        try:
            compte_rendu = reseau.verifier_acces(reseau.creer_session(mode))
            dep.sortie(f"  mode {mode:8}: OK ({compte_rendu})")
            ok = True
        except Exception as exc:  # on veut voir toutes les erreurs ici
            premiere_ligne = str(exc).splitlines()[0] if str(exc) else type(exc).__name__
            dep.sortie(f"  mode {mode:8}: ÉCHEC - {premiere_ligne[:300]}")
    return 0 if ok else 1


@dataclass
class Contexte:
    args: argparse.Namespace
    config: Config
    regles: Regles
    source: str
    dep: Dependances
    registre: Registre


def _oui(reponse: str) -> bool:
    return reponse.strip().lower() in {"o", "oui", "y", "yes"}


def _traiter(lien: str, cible: CibleSondage, ctx: Contexte, nom: str, actualisation: bool = False) -> int:
    """Lit le sondage et l'agenda, affiche l'aperçu et l'écart, puis envoie si demandé."""
    dep, args = ctx.dep, ctx.args
    assert dep.fabrique_source is not None
    dep.sortie(f"Lecture du sondage {cible.nom}...")
    sondage = cible.lire(lien)
    jours = sorted({dep.fuseau.vers_local(t).date() for t in sondage.creneaux})
    dep.sortie(f"Lecture de l'agenda Outlook ({len(jours)} jour(s))...")
    grille = dep.fabrique_source(dep.fuseau, ctx.source).grille(jours)
    creneaux = evaluer(sondage.creneaux, grille, ctx.regles, dep.fuseau, sondage.duree_creneau)

    avant = sondage.participants.get(nom) if nom else None
    if avant is not None:
        # Les créneaux déjà passés restent tels qu'en ligne : inutile de les réécrire.
        maintenant = dep.maintenant()
        creneaux = [
            replace(c, libre=i in avant, motif=Motif.PASSE)
            if c.epoch + sondage.duree_creneau <= maintenant else c
            for i, c in enumerate(creneaux)
        ]

    dep.sortie("")
    dep.sortie(formater(sondage, creneaux, ctx.regles, grille.nb_rendez_vous, detail=not actualisation))
    if nom:
        dep.sortie("")
        dep.sortie(formater_ecart(creneaux, avant, nom))
    dep.sortie("")

    if not args.submit:
        dep.sortie("Aperçu seulement : rien n'a été envoyé (ajoutez --submit pour envoyer).")
        return 0

    while not nom:
        nom = dep.entree("Votre nom dans le sondage : ").strip()
        avant = sondage.participants.get(nom)
        if nom:
            dep.sortie(formater_ecart(creneaux, avant, nom))
    disponibilites = [c.libre for c in creneaux]

    if avant is not None and {i for i, d in enumerate(disponibilites) if d} == set(avant):
        ctx.registre.enregistrer(sondage.url, sondage.titre, nom, sondage.fin, dep.maintenant())
        dep.sortie("Rien à envoyer.")
        return 0

    mot_de_passe = args.mot_de_passe if args.mot_de_passe is not None else ctx.config.mot_de_passe
    if not args.oui:
        dep.sortie(f"Participant : {nom}")
        if not _oui(dep.entree("Envoyer ? (O/N) ")):
            dep.sortie("Envoi annulé : rien n'a été envoyé.")
            return 0

    dep.sortie(f"Envoi à {cible.nom}...")
    resultat = cible.envoyer(sondage, nom, mot_de_passe, disponibilites)
    ctx.registre.enregistrer(sondage.url, sondage.titre, nom, sondage.fin, dep.maintenant())
    rappel = "Dans le navigateur, rechargez la page du sondage (F5) pour voir la mise à jour."
    if resultat.verifie:
        dep.sortie(f"OK : {resultat.nb_libres} créneau(x) libres enregistrés et vérifiés sur la page.")
        dep.sortie(rappel)
        return 0
    if resultat.verifie is None:
        dep.sortie(f"Envoyé, mais non vérifié ({resultat.detail}). Contrôlez le sondage dans le navigateur.")
        dep.sortie(rappel)
        return 0
    dep.sortie(
        "ATTENTION : l'envoi a été accepté mais la page ne reflète pas ce qui a été envoyé "
        f"({resultat.detail}). Contrôlez le sondage dans le navigateur."
    )
    return 3


def _purger(ctx: Contexte) -> None:
    for e in ctx.registre.purger_termines(ctx.dep.maintenant()):
        ctx.dep.sortie(f"Sondage terminé, retiré de la liste : {e.titre}")


def _liste(ctx: Contexte) -> int:
    _purger(ctx)
    entrees = ctx.registre.lire()
    if not entrees:
        ctx.dep.sortie("Aucun sondage en cours de suivi.")
        return 0
    ctx.dep.sortie("Sondages suivis (actualisés par o2s-actualiser.bat) :")
    for n, e in enumerate(entrees, 1):
        fin = ctx.dep.fuseau.vers_local(e.fin)
        ctx.dep.sortie(f"  {n}. {e.titre}  [{e.nom}]  jusqu'au {fin:%d/%m/%Y}, dernier envoi {e.dernier_envoi}")
        ctx.dep.sortie(f"     {e.url}")
    return 0


def _oublier(ctx: Contexte, cibles: list[CibleSondage], quoi: str) -> int:
    entrees = ctx.registre.lire()
    if not entrees:
        ctx.dep.sortie("Aucun sondage en cours de suivi.")
        return 0
    if not quoi:
        _liste(ctx)
        quoi = ctx.dep.entree("Numéro (ou lien) du sondage à ne plus actualiser : ").strip()
    if quoi.isdigit() and 1 <= int(quoi) <= len(entrees):
        url = entrees[int(quoi) - 1].url
    else:
        cible = _choisir_cible(quoi, cibles)
        if cible is None:
            raise ErreurUtilisateur(f"Ni un numéro de la liste ni un lien de sondage : {quoi!r}.")
        url = cible.normaliser(quoi)
    retire = ctx.registre.oublier(url)
    if retire is None:
        ctx.dep.sortie("Ce sondage n'était pas dans la liste.")
    else:
        ctx.dep.sortie(f"Ne sera plus actualisé : {retire.titre}. (Vos disponibilités en ligne restent telles quelles.)")
    return 0


def _actualiser(ctx: Contexte, cibles: list[CibleSondage]) -> int:
    _purger(ctx)
    entrees = ctx.registre.lire()
    if not entrees:
        ctx.dep.sortie("Aucun sondage à actualiser : envoyez d'abord un sondage avec o2s.bat.")
        return 0
    ctx.dep.sortie(
        f"{len(entrees)} sondage(s) à actualiser. Rappel : une fois la date choisie, retirez le "
        "sondage de la liste (o2s.bat --oublier) avant d'ajouter la réunion à votre agenda."
    )
    pire = 0
    for n, e in enumerate(entrees, 1):
        ctx.dep.sortie("")
        ctx.dep.sortie(f"===== [{n}/{len(entrees)}] {e.titre} =====")
        cible = _choisir_cible(e.url, cibles)
        if cible is None:
            ctx.dep.sortie(f"Outil de sondage non pris en charge : {e.url}")
            continue
        try:
            pire = max(pire, _traiter(e.url, cible, ctx, e.nom, actualisation=True))
        except ErreurUtilisateur as exc:
            ctx.dep.sortie(f"ERREUR : {exc}")
            pire = max(pire, 1)
    return pire


def executer(argv: list[str] | None, dep: Dependances | None = None) -> int:
    dep = dep or Dependances()
    dep.fabrique_source = dep.fabrique_source or _source_outlook
    dep.fabrique_cibles = dep.fabrique_cibles or _cibles_par_defaut
    args = construire_parser().parse_args(argv)
    config = charger(args.config)
    regles = regles_depuis(config, args)

    mode_reseau = args.reseau or config.reseau or reseau.mode_par_defaut()
    if mode_reseau not in reseau.MODES:
        raise ErreurUtilisateur(f"Mode réseau inconnu dans la configuration : {mode_reseau!r} (windows ou python).")

    source = args.source or config.source_agenda
    if source not in SOURCES:
        raise ErreurUtilisateur(f"Source d'agenda inconnue : {source!r} (rendez-vous ou freebusy).")

    if args.test_reseau:
        return _diagnostic_reseau(dep)
    if args.agenda:
        return _diagnostic_agenda(args.agenda, dep, source)

    ctx = Contexte(args, config, regles, source, dep, Registre())
    if args.liste:
        return _liste(ctx)
    cibles = dep.fabrique_cibles(mode_reseau)
    if args.oublier is not None:
        return _oublier(ctx, cibles, args.oublier.strip())
    if args.actualiser:
        return _actualiser(ctx, cibles)

    lien, cible = _obtenir_lien(args, cibles, dep)
    return _traiter(lien, cible, ctx, (args.nom or config.nom).strip())


def main(argv: list[str] | None = None) -> int:
    for flux in (sys.stdout, sys.stderr):
        try:
            flux.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass
    try:
        return executer(argv)
    except ErreurUtilisateur as exc:
        print(f"\nERREUR : {exc}", file=sys.stderr)
        return 1
    except (KeyboardInterrupt, EOFError):
        print("\nAnnulé.", file=sys.stderr)
        return 130
    except Exception:  # pragma: no cover
        traceback.print_exc()
        print("\nERREUR inattendue (voir le détail ci-dessus).", file=sys.stderr)
        return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
