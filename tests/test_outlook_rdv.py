from datetime import date, datetime, timedelta, timezone

import pytest

from outlook2scheduler.erreurs import ErreurSecuriteOutlook
from outlook2scheduler.sources.base import AILLEURS, LIBRE, OCCUPE, PROVISOIRE
from outlook2scheduler.sources.outlook import DISP_E_EXCEPTION, E_ABORT
from outlook2scheduler.sources.outlook_rdv import (
    RendezVous,
    SourceOutlookRendezVous,
    filtre_restrict,
    lire_elements,
)

PROPRIETES_INTERDITES = {"Subject", "Body", "Organizer", "Recipients", "RequiredAttendees",
                         "OptionalAttendees", "Location", "Categories", "HTMLBody", "RTFBody"}


class FausseComError(Exception):
    pass


class FauxElement:
    """Rendez-vous Outlook. Lire une propriété de contenu fait échouer le test."""

    def __init__(self, debut_utc, fin_utc, statut=OCCUPE, meeting=0, reponse=0, tz=None):
        self.StartUTC = debut_utc.replace(tzinfo=tz)
        self.EndUTC = fin_utc.replace(tzinfo=tz)
        self.BusyStatus = statut
        self.MeetingStatus = meeting
        self.ResponseStatus = reponse

    def __getattribute__(self, nom):
        if nom in PROPRIETES_INTERDITES:
            raise AssertionError(f"propriété confidentielle lue : {nom}")
        return object.__getattribute__(self, nom)


class FauxItems:
    def __init__(self, elements, erreur=None):
        self._elements, self._erreur = elements, erreur
        self.journal = []
        self._pos = 0

    def Sort(self, cle):
        self.journal.append(("Sort", cle))

    def __setattr__(self, nom, valeur):
        if nom == "IncludeRecurrences":
            self.journal.append(("IncludeRecurrences", valeur))
        object.__setattr__(self, nom, valeur)

    def Restrict(self, filtre):
        if self._erreur:
            raise self._erreur
        self.journal.append(("Restrict", filtre))
        return self

    def GetFirst(self):
        self._pos = 0
        return self.GetNext()

    def GetNext(self):
        if self._pos >= len(self._elements):
            return None
        self._pos += 1
        return self._elements[self._pos - 1]


class FauxDossier:
    def __init__(self, items):
        self.Items = items


def iso(dt):
    return dt.strftime("%Y-%m-%d %H:%M")


def test_ordre_des_appels_et_filtre():
    items = FauxItems([])
    lire_elements(FauxDossier(items), datetime(2026, 11, 1), datetime(2026, 11, 4), iso)
    assert items.journal == [
        ("Sort", "[Start]"),
        ("IncludeRecurrences", True),
        ("Restrict", "[Start] < '2026-11-04 00:00' AND [End] > '2026-11-01 00:00'"),
    ]
    assert filtre_restrict(datetime(2026, 1, 2), datetime(2026, 1, 3), iso).startswith("[Start] < '2026-01-03")


def test_statuts_filtres_et_aucune_propriete_confidentielle_lue():
    h = datetime(2026, 11, 2, 9)
    une_heure = timedelta(hours=1)
    elements = [
        FauxElement(h, h + une_heure, OCCUPE),
        FauxElement(h, h + une_heure, LIBRE),
        FauxElement(h, h + une_heure, OCCUPE, meeting=5),  # annulée
        FauxElement(h, h + une_heure, PROVISOIRE, reponse=4),  # refusée
        FauxElement(h, h + une_heure, PROVISOIRE),
        FauxElement(h, h, OCCUPE),  # durée nulle
    ]
    rdvs = lire_elements(FauxDossier(FauxItems(elements)), h, h + une_heure, iso)
    assert [r.statut for r in rdvs] == [OCCUPE, PROVISOIRE]
    assert rdvs[0].debut == int(h.replace(tzinfo=timezone.utc).timestamp())


def test_fuseau_trompeur_de_pywin32_ignore():
    trompeur = timezone(timedelta(hours=5))
    h = datetime(2026, 11, 2, 9)
    rdv = lire_elements(FauxDossier(FauxItems([FauxElement(h, h + timedelta(hours=1), tz=trompeur)])), h, h, iso)[0]
    assert rdv.debut == int(h.replace(tzinfo=timezone.utc).timestamp())


def test_refus_de_securite_traduit():
    exc = FausseComError(DISP_E_EXCEPTION, "", (0, "", "", None, 0, E_ABORT), None)
    with pytest.raises(ErreurSecuriteOutlook):
        lire_elements(FauxDossier(FauxItems([], erreur=exc)), datetime(2026, 1, 1), datetime(2026, 1, 2), iso, FausseComError)


class FauxClient:
    def __init__(self, fuseau, rdvs_locaux):
        self.fuseau = fuseau
        self.rdvs = [RendezVous(fuseau.vers_epoch(d), fuseau.vers_epoch(f), s) for d, f, s in rdvs_locaux]
        self.fenetres = []

    def rendez_vous(self, debut, fin):
        self.fenetres.append((debut, fin))
        return self.rdvs


def code(grille, fuseau, *dt):
    return grille.code(fuseau.vers_epoch(datetime(*dt)))


def test_grille_chevauchements_et_priorite(fuseau):
    client = FauxClient(fuseau, [
        (datetime(2026, 11, 2, 10), datetime(2026, 11, 2, 12), AILLEURS),
        (datetime(2026, 11, 2, 11), datetime(2026, 11, 2, 11, 30), OCCUPE),
        (datetime(2026, 11, 2, 23, 30), datetime(2026, 11, 3, 0, 30), PROVISOIRE),  # passe minuit
        (datetime(2026, 11, 9, 10), datetime(2026, 11, 9, 11), OCCUPE),  # hors des jours demandés
    ])
    grille = SourceOutlookRendezVous(client, fuseau).grille([date(2026, 11, 2), date(2026, 11, 3)])
    assert code(grille, fuseau, 2026, 11, 2, 9, 45) == LIBRE
    assert code(grille, fuseau, 2026, 11, 2, 10, 0) == AILLEURS
    assert code(grille, fuseau, 2026, 11, 2, 11, 15) == OCCUPE
    assert code(grille, fuseau, 2026, 11, 2, 11, 30) == AILLEURS
    assert code(grille, fuseau, 2026, 11, 3, 0, 15) == PROVISOIRE
    assert code(grille, fuseau, 2026, 11, 4, 9, 0) is None
    assert grille.nb_rendez_vous == 3
    assert client.fenetres == [(datetime(2026, 11, 1), datetime(2026, 11, 5))]


def test_grille_juste_le_jour_du_changement_d_heure(fuseau):
    # Dimanche 25/10/2026 : 3h devient 2h. Les heures UTC lèvent toute ambiguïté.
    client = FauxClient(fuseau, [(datetime(2026, 10, 25, 10), datetime(2026, 10, 25, 11), OCCUPE)])
    grille = SourceOutlookRendezVous(client, fuseau).grille([date(2026, 10, 25)])
    assert len([q for q in grille.codes]) == 100  # journée de 25 h
    assert code(grille, fuseau, 2026, 10, 25, 9, 45) == LIBRE
    assert code(grille, fuseau, 2026, 10, 25, 10, 0) == OCCUPE
    assert code(grille, fuseau, 2026, 10, 25, 11, 0) == LIBRE
    assert grille.jours_changement_heure == set()
