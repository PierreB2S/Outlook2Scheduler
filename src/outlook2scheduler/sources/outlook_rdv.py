"""Lecture de l'agenda Outlook élément par élément (source par défaut).

Pourquoi : ``CurrentUser.FreeBusy`` passe par des objets que la protection
d'Outlook (« Un programme tente d'accéder... ») bloque sur certains postes.
Ici on ne touche qu'à des propriétés non protégées du calendrier par défaut :
``StartUTC``, ``EndUTC``, ``BusyStatus``, ``MeetingStatus``, ``ResponseStatus``.
L'objet, les participants et le contenu ne sont jamais lus.

Les heures sont lues en UTC (StartUTC/EndUTC) : pas d'ambiguïté au
changement d'heure.
"""

from __future__ import annotations

import calendar
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Callable, Iterable, Protocol

from outlook2scheduler.sources.base import (
    ABSENT,
    AILLEURS,
    LIBRE,
    OCCUPE,
    PAS_SECONDES,
    PROVISOIRE,
    GrilleAgenda,
)
from outlook2scheduler.sources.outlook import connecter_outlook, traduire_erreur_com
from outlook2scheduler.temps import Fuseau, minuit

OL_DOSSIER_CALENDRIER = 9
REUNIONS_ANNULEES = {5, 7}  # olMeetingCanceled, olMeetingReceivedAndCanceled
REPONSE_REFUSEE = 4  # olResponseDeclined
LIMITE_ELEMENTS = 50_000

# Quand deux rendez-vous se chevauchent, le statut le plus contraignant l'emporte.
PRIORITE = {LIBRE: 0, AILLEURS: 1, PROVISOIRE: 2, OCCUPE: 3, ABSENT: 4}


@dataclass(frozen=True)
class RendezVous:
    debut: int  # epoch
    fin: int  # epoch
    statut: int  # code BusyStatus (0 à 4)


class ClientRendezVous(Protocol):
    def rendez_vous(self, debut: datetime, fin: datetime) -> list[RendezVous]:
        """Rendez-vous (heures locales naïves) chevauchant [debut, fin)."""


def format_date_systeme(dt: datetime) -> str:
    """Date au format court de Windows + HH:MM, comme le Format(…, "ddddd") des exemples VBA.

    Items.Restrict interprète les dates selon les paramètres régionaux du poste.
    """
    if sys.platform == "win32":  # pragma: no cover - dépend de Windows
        import ctypes
        from ctypes import wintypes

        class SYSTEMTIME(ctypes.Structure):
            _fields_ = [(n, wintypes.WORD) for n in
                        ("wYear", "wMonth", "wDayOfWeek", "wDay", "wHour", "wMinute", "wSecond", "wMilliseconds")]

        st = SYSTEMTIME(dt.year, dt.month, 0, dt.day, 0, 0, 0, 0)
        tampon = ctypes.create_unicode_buffer(64)
        LOCALE_USER_DEFAULT, DATE_SHORTDATE = 0x0400, 0x1
        if ctypes.windll.kernel32.GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, ctypes.byref(st), None, tampon, 64):
            return f"{tampon.value} {dt:%H:%M}"
    return f"{dt:%m/%d/%Y %H:%M}"


def filtre_restrict(debut: datetime, fin: datetime, formater: Callable[[datetime], str]) -> str:
    return f"[Start] < '{formater(fin)}' AND [End] > '{formater(debut)}'"


def _epoch_utc(valeur: Any) -> int:
    """Valeur de StartUTC/EndUTC (datetime pywin32) -> epoch.

    pywin32 peut accrocher un fuseau trompeur aux dates COM : on ne garde que
    les champs bruts, qui sont bien en UTC pour StartUTC/EndUTC.
    """
    return calendar.timegm(valeur.replace(tzinfo=None).timetuple())


def lire_elements(
    dossier: Any,
    debut: datetime,
    fin: datetime,
    formater: Callable[[datetime], str],
    com_error: type[BaseException] = Exception,
) -> list[RendezVous]:
    """Parcourt les éléments du calendrier (récurrences développées) entre debut et fin."""
    try:
        elements = dossier.Items
        elements.Sort("[Start]")
        elements.IncludeRecurrences = True
        restreints = elements.Restrict(filtre_restrict(debut, fin, formater))
        element = restreints.GetFirst()
    except com_error as exc:
        raise traduire_erreur_com(exc, "lecture du calendrier") from exc

    resultat: list[RendezVous] = []
    for _ in range(LIMITE_ELEMENTS):
        if element is None:
            break
        try:
            statut = int(element.BusyStatus)
            ignorer = (
                statut == LIBRE
                or int(getattr(element, "MeetingStatus", 0)) in REUNIONS_ANNULEES
                or int(getattr(element, "ResponseStatus", 0)) == REPONSE_REFUSEE
            )
            if not ignorer:
                d, f = _epoch_utc(element.StartUTC), _epoch_utc(element.EndUTC)
                if f > d:
                    resultat.append(RendezVous(d, f, statut))
        except (com_error, AttributeError, TypeError, ValueError):
            pass  # élément d'un autre type ou illisible : ignoré
        try:
            element = restreints.GetNext()
        except com_error as exc:
            raise traduire_erreur_com(exc, "lecture du calendrier") from exc
    return resultat


class OutlookRendezVousCOM:  # pragma: no cover - nécessite Outlook
    def __init__(self) -> None:
        app, self._com_error = connecter_outlook()
        try:
            self._dossier = app.GetNamespace("MAPI").GetDefaultFolder(OL_DOSSIER_CALENDRIER)
        except self._com_error as exc:
            raise traduire_erreur_com(exc, "ouverture du calendrier") from exc

    def rendez_vous(self, debut: datetime, fin: datetime) -> list[RendezVous]:
        return lire_elements(self._dossier, debut, fin, format_date_systeme, self._com_error)


class SourceOutlookRendezVous:
    def __init__(self, client: ClientRendezVous, fuseau: Fuseau) -> None:
        self.client = client
        self.fuseau = fuseau

    def grille(self, jours: Iterable[date]) -> GrilleAgenda:
        jours = sorted(set(jours))
        grille = GrilleAgenda()
        if not jours:
            grille.nb_rendez_vous = 0
            return grille
        for jour in jours:
            q = self.fuseau.vers_epoch(minuit(jour))
            fin_jour = self.fuseau.vers_epoch(minuit(jour + timedelta(days=1)))
            while q < fin_jour:
                grille.codes[q] = LIBRE
                q += PAS_SECONDES

        # Fenêtre élargie d'un jour de chaque côté : le tri fin se fait ici, en UTC.
        rdvs = self.client.rendez_vous(minuit(jours[0] - timedelta(days=1)), minuit(jours[-1] + timedelta(days=2)))
        grille.nb_rendez_vous = 0
        for rdv in rdvs:
            touche = False
            q = rdv.debut - rdv.debut % PAS_SECONDES
            while q < rdv.fin:
                actuel = grille.codes.get(q)
                if actuel is not None:
                    touche = True
                    if PRIORITE.get(rdv.statut, 4) > PRIORITE.get(actuel, 0):
                        grille.codes[q] = rdv.statut
                q += PAS_SECONDES
            grille.nb_rendez_vous += touche
        return grille
