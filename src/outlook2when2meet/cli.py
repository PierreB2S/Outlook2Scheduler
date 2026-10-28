"""Ligne de commande : w2m <lien> [--nom] [--heures 08:30-18:00] [--provisoire-libre] [--submit]."""

from __future__ import annotations

import argparse
import sys
import traceback
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Callable

from outlook2when2meet import __version__
from outlook2when2meet.apercu import formater
from outlook2when2meet.cibles.base import CibleSondage
from outlook2when2meet.config import Config, charger
from outlook2when2meet.erreurs import ErreurUtilisateur
from outlook2when2meet.regles import Regles, evaluer, parse_heures, parse_jours
from outlook2when2meet.sources.base import PAS_SECONDES, SourceAgenda
from outlook2when2meet.temps import JOURS_COURTS, Fuseau, FuseauSysteme, minuit

CODES_LISIBLES = {"0": ".", "1": "p", "2": "X", "3": "A", "4": "a"}


def creer_session():  # pragma: no cover - dépend du réseau du poste
    """Session HTTP utilisant les certificats Windows (proxy avec inspection SSL)."""
    import requests
    import truststore

    truststore.inject_into_ssl()
    session = requests.Session()
    session.headers["User-Agent"] = f"outlook2when2meet/{__version__}"
    return session


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

    fabrique_source: Callable[[Fuseau], SourceAgenda] | None = None
    fabrique_cibles: Callable[[], list[CibleSondage]] | None = None
    fuseau: Fuseau = field(default_factory=FuseauSysteme)
    entree: Callable[[str], str] = input
    sortie: Callable[[str], None] = print
    presse_papiers: Callable[[], str] = lire_presse_papiers


def _source_outlook(fuseau: Fuseau) -> SourceAgenda:  # pragma: no cover
    from outlook2when2meet.sources.outlook import OutlookCOM, SourceOutlook

    return SourceOutlook(OutlookCOM(), fuseau)


def _cibles_par_defaut() -> list[CibleSondage]:  # pragma: no cover
    from outlook2when2meet.cibles.when2meet import When2meet

    return [When2meet(creer_session())]


def construire_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="w2m",
        description="Remplit un sondage When2meet à partir de votre agenda Outlook. "
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
        "--agenda",
        metavar="AAAA-MM-JJ",
        help="diagnostic : affiche la grille Outlook de ce jour (rien n'est envoyé)",
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


def _diagnostic_agenda(jour_texte: str, dep: Dependances) -> int:
    try:
        jour = date.fromisoformat(jour_texte)
    except ValueError as exc:
        raise ErreurUtilisateur(f"Date invalide : {jour_texte!r} (format AAAA-MM-JJ).") from exc
    assert dep.fabrique_source is not None
    grille = dep.fabrique_source(dep.fuseau).grille([jour])
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
    if jour in grille.jours_changement_heure:
        dep.sortie("Attention : jour de changement d'heure.")
    return 0


def executer(argv: list[str] | None, dep: Dependances | None = None) -> int:
    dep = dep or Dependances()
    dep.fabrique_source = dep.fabrique_source or _source_outlook
    dep.fabrique_cibles = dep.fabrique_cibles or _cibles_par_defaut
    args = construire_parser().parse_args(argv)
    config = charger(args.config)
    regles = regles_depuis(config, args)

    if args.agenda:
        return _diagnostic_agenda(args.agenda, dep)

    cibles = dep.fabrique_cibles()
    lien, cible = _obtenir_lien(args, cibles, dep)

    dep.sortie(f"Lecture du sondage {cible.nom}...")
    sondage = cible.lire(lien)
    jours = sorted({dep.fuseau.vers_local(t).date() for t in sondage.creneaux})
    dep.sortie(f"Lecture de l'agenda Outlook ({len(jours)} jour(s))...")
    grille = dep.fabrique_source(dep.fuseau).grille(jours)
    creneaux = evaluer(sondage.creneaux, grille, regles, dep.fuseau, sondage.duree_creneau)

    dep.sortie("")
    dep.sortie(formater(sondage, creneaux, regles))
    dep.sortie("")

    if not args.submit:
        dep.sortie("Aperçu seulement : rien n'a été envoyé (ajoutez --submit pour envoyer).")
        return 0

    nom = (args.nom or config.nom).strip()
    while not nom:
        nom = dep.entree("Votre nom dans le sondage : ").strip()
    mot_de_passe = args.mot_de_passe if args.mot_de_passe is not None else config.mot_de_passe

    if not args.oui:
        dep.sortie(f"Participant : {nom}")
        reponse = dep.entree("Envoyer ? (O/N) ")
        if reponse.strip().lower() not in {"o", "oui", "y", "yes"}:
            dep.sortie("Envoi annulé : rien n'a été envoyé.")
            return 0

    dep.sortie(f"Envoi à {cible.nom}...")
    resultat = cible.envoyer(sondage, nom, mot_de_passe, [c.libre for c in creneaux])
    if resultat.verifie:
        dep.sortie(f"OK : {resultat.nb_libres} créneau(x) libres enregistrés et vérifiés sur la page.")
        return 0
    if resultat.verifie is None:
        dep.sortie(f"Envoyé, mais non vérifié ({resultat.detail}). Contrôlez le sondage dans le navigateur.")
        return 0
    dep.sortie(
        "ATTENTION : l'envoi a été accepté mais la page ne reflète pas ce qui a été envoyé "
        f"({resultat.detail}). Contrôlez le sondage dans le navigateur."
    )
    return 3


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
