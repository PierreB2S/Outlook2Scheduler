"""Lecture de l'agenda Outlook classique via COM (pywin32).

On n'utilise que ``Recipient.FreeBusy(date, 15, True)`` sur l'utilisateur
courant : Outlook renvoie une chaîne de chiffres (un par quart d'heure) à
partir de minuit du jour demandé. Aucun élément d'agenda (intitulé,
participants...) n'est ouvert.

Changement d'heure : on ne sait pas avec certitude si Outlook compte les
quarts d'heure en temps écoulé ou en heure murale le jour du passage. Pour
que ça n'ait pas d'importance, les blocs demandés à Outlook ne chevauchent
jamais un changement d'heure : le jour du changement est lu seul et signalé.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Iterable, Protocol

from outlook2scheduler.erreurs import (
    ErreurOutlook,
    ErreurOutlookFerme,
    ErreurSecuriteOutlook,
)
from outlook2scheduler.sources.base import PAS_SECONDES, GrilleAgenda
from outlook2scheduler.temps import Fuseau, est_jour_changement_heure, minuit

TAILLE_BLOC_JOURS = 7

# HRESULT utiles (valeurs signées telles que renvoyées par pywin32)
MK_E_UNAVAILABLE = -2147221021  # 0x800401E3 : pas d'Outlook en cours d'exécution
CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005
E_ABORT = -2147467260  # 0x80004004 : alerte de sécurité refusée
E_ACCESSDENIED = -2147024891  # 0x80070005
DISP_E_EXCEPTION = -2147352567  # 0x80020009 : erreur interne, vrai code dans excepinfo


class ClientFreeBusy(Protocol):
    def free_busy(self, debut: datetime, pas_minutes: int) -> str:
        """Chaîne FreeBusy commençant à minuit du jour de ``debut``."""


def _hresults(exc: BaseException) -> set[int]:
    """Extrait les HRESULT d'une com_error (code principal et code interne)."""
    codes: set[int] = set()
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        codes.add(args[0])
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 5:
        scode = args[2][5]
        if isinstance(scode, int):
            codes.add(scode)
    return codes


def traduire_erreur_com(exc: BaseException, etape: str) -> ErreurOutlook:
    codes = _hresults(exc)
    if codes & {MK_E_UNAVAILABLE, CO_E_SERVER_EXEC_FAILURE} and etape == "connexion":
        return ErreurOutlookFerme(
            "Outlook ne semble pas ouvert. Lancez Outlook (version classique), "
            "attendez qu'il ait fini de démarrer, puis relancez.\n"
            "Si Outlook est bien ouvert : il tourne peut-être « en tant "
            "qu'administrateur » et pas ce script (ou l'inverse) ; lancez les "
            "deux de la même façon."
        )
    if codes & {E_ABORT, E_ACCESSDENIED}:
        return ErreurSecuriteOutlook(
            "Outlook a refusé l'accès à l'agenda (alerte de sécurité « Un "
            "programme tente d'accéder... »).\n"
            "Relancez et cliquez sur « Autoriser » dans la fenêtre d'Outlook "
            "(elle peut être cachée derrière les autres fenêtres). Si l'alerte "
            "revient à chaque fois ou si l'accès est bloqué sans alerte, la "
            "politique de sécurité du poste interdit l'accès par programme : "
            "voir le README (section Limites)."
        )
    detail = " / ".join(f"0x{c & 0xFFFFFFFF:08X}" for c in sorted(codes)) or str(exc)
    return ErreurOutlook(
        f"Erreur Outlook pendant l'étape « {etape} » ({detail}).\n"
        "Vérifiez qu'Outlook est ouvert, connecté (pas en mode hors connexion) "
        "et qu'aucune fenêtre modale n'y est ouverte, puis relancez."
    )


class OutlookCOM:
    """Accès réel à Outlook par COM. Non testable hors d'un poste Windows + Outlook."""

    def __init__(self) -> None:
        try:
            import pywintypes  # noqa: F401
            import win32com.client
        except ImportError as exc:  # pragma: no cover - dépend du poste
            raise ErreurOutlook(
                "pywin32 est introuvable : cet outil ne fonctionne que sous "
                "Windows, lancé via « uv run » (ou w2m.bat)."
            ) from exc
        self._com_error = pywintypes.com_error
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except self._com_error as exc:
            raise traduire_erreur_com(exc, "connexion") from exc
        try:
            self._utilisateur = app.GetNamespace("MAPI").CurrentUser
        except self._com_error as exc:
            raise traduire_erreur_com(exc, "accès au profil") from exc

    def free_busy(self, debut: datetime, pas_minutes: int) -> str:  # pragma: no cover
        try:
            resultat = self._utilisateur.FreeBusy(debut, pas_minutes, True)
        except self._com_error as exc:
            raise traduire_erreur_com(exc, "lecture libre/occupé") from exc
        return str(resultat or "")


def decouper_blocs(
    jours: Iterable[date], fuseau: Fuseau, taille: int = TAILLE_BLOC_JOURS
) -> list[list[date]]:
    """Regroupe les jours en blocs consécutifs d'au plus ``taille`` jours.

    Un jour de changement d'heure forme toujours un bloc à lui seul, pour
    qu'aucun bloc ne contienne de changement d'heure après son premier jour.
    """
    blocs: list[list[date]] = []
    courant: list[date] = []
    for jour in sorted(set(jours)):
        special = est_jour_changement_heure(fuseau, jour)
        continue_bloc = (
            courant
            and not special
            and not est_jour_changement_heure(fuseau, courant[-1])
            and jour == courant[-1] + timedelta(days=1)
            and len(courant) < taille
        )
        if continue_bloc:
            courant.append(jour)
        else:
            if courant:
                blocs.append(courant)
            courant = [jour]
    if courant:
        blocs.append(courant)
    return blocs


class SourceOutlook:
    def __init__(self, client: ClientFreeBusy, fuseau: Fuseau) -> None:
        self.client = client
        self.fuseau = fuseau

    def grille(self, jours: Iterable[date]) -> GrilleAgenda:
        grille = GrilleAgenda()
        for bloc in decouper_blocs(jours, self.fuseau):
            debut = minuit(bloc[0])
            fin = minuit(bloc[-1] + timedelta(days=1))
            base = self.fuseau.vers_epoch(debut)
            nb_quarts = (self.fuseau.vers_epoch(fin) - base) // PAS_SECONDES
            chaine = self.client.free_busy(debut, PAS_SECONDES // 60)
            for i, car in enumerate(chaine[:nb_quarts]):
                if car.isdigit():
                    grille.codes[base + i * PAS_SECONDES] = int(car)
            for jour in bloc:
                if est_jour_changement_heure(self.fuseau, jour):
                    grille.jours_changement_heure.add(jour)
        return grille
