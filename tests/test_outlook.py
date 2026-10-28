from datetime import date, datetime

import pytest
from conftest import FauxFreeBusy

from outlook2when2meet.erreurs import ErreurOutlook, ErreurOutlookFerme, ErreurSecuriteOutlook
from outlook2when2meet.sources.base import LIBRE, OCCUPE, PROVISOIRE
from outlook2when2meet.sources.outlook import (
    DISP_E_EXCEPTION,
    E_ABORT,
    MK_E_UNAVAILABLE,
    SourceOutlook,
    decouper_blocs,
    traduire_erreur_com,
)
from outlook2when2meet.temps import est_jour_changement_heure

# Passage à l'heure d'hiver : dimanche 25/10/2026 ; à l'heure d'été : dimanche 28/03/2027.
HIVER = date(2026, 10, 25)
ETE = date(2027, 3, 28)


def jours(debut: date, n: int) -> list[date]:
    return [date.fromordinal(debut.toordinal() + i) for i in range(n)]


def test_detection_changement_heure(fuseau):
    assert est_jour_changement_heure(fuseau, HIVER)
    assert est_jour_changement_heure(fuseau, ETE)
    assert not est_jour_changement_heure(fuseau, date(2026, 10, 26))
    assert not est_jour_changement_heure(fuseau, date(2026, 11, 3))


def test_blocs_de_sept_jours_au_plus(fuseau):
    blocs = decouper_blocs(jours(date(2026, 11, 2), 10), fuseau)
    assert [len(b) for b in blocs] == [7, 3]


def test_blocs_dates_non_consecutives(fuseau):
    d = [date(2026, 11, 2), date(2026, 11, 3), date(2026, 11, 10), date(2026, 12, 1)]
    assert decouper_blocs(d, fuseau) == [d[:2], [d[2]], [d[3]]]


def test_le_jour_du_changement_d_heure_est_isole(fuseau):
    blocs = decouper_blocs(jours(date(2026, 10, 22), 7), fuseau)
    assert blocs == [
        [date(2026, 10, 22), date(2026, 10, 23), date(2026, 10, 24)],
        [HIVER],
        [date(2026, 10, 26), date(2026, 10, 27), date(2026, 10, 28)],
    ]


@pytest.mark.parametrize("semantique", ["ecoule", "murale"])
@pytest.mark.parametrize("veille,lendemain", [(date(2026, 10, 23), date(2026, 10, 26)), (date(2027, 3, 26), date(2027, 3, 29))])
def test_grille_juste_de_part_et_d_autre_du_changement_d_heure(fuseau, semantique, veille, lendemain):
    evenements = [
        (datetime.combine(veille, datetime.min.time()).replace(hour=10), datetime.combine(veille, datetime.min.time()).replace(hour=11), OCCUPE),
        (datetime.combine(lendemain, datetime.min.time()).replace(hour=10), datetime.combine(lendemain, datetime.min.time()).replace(hour=11), PROVISOIRE),
    ]
    client = FauxFreeBusy(fuseau, evenements, semantique=semantique)
    grille = SourceOutlook(client, fuseau).grille(jours(veille, 4))

    def code(jour, h, m=0):
        return grille.code(fuseau.vers_epoch(datetime(jour.year, jour.month, jour.day, h, m)))

    assert code(veille, 9, 45) == LIBRE
    assert code(veille, 10, 0) == OCCUPE
    assert code(veille, 10, 45) == OCCUPE
    assert code(veille, 11, 0) == LIBRE
    assert code(lendemain, 9, 45) == LIBRE
    assert code(lendemain, 10, 0) == PROVISOIRE
    assert code(lendemain, 10, 45) == PROVISOIRE
    assert code(lendemain, 11, 0) == LIBRE
    assert len(grille.jours_changement_heure) == 1


def test_appels_par_bloc_a_minuit(fuseau):
    client = FauxFreeBusy(fuseau)
    SourceOutlook(client, fuseau).grille(jours(date(2026, 11, 2), 10))
    assert client.appels == [datetime(2026, 11, 2), datetime(2026, 11, 9)]


def test_hors_de_la_grille_outlook(fuseau):
    # Outlook ne renvoie que 2 jours : le 3e jour du bloc est inconnu.
    client = FauxFreeBusy(fuseau, horizon_jours=2)
    grille = SourceOutlook(client, fuseau).grille(jours(date(2026, 11, 2), 3))
    assert grille.code(fuseau.vers_epoch(datetime(2026, 11, 3, 23, 45))) == LIBRE
    assert grille.code(fuseau.vers_epoch(datetime(2026, 11, 4, 9, 0))) is None


def test_caracteres_non_numeriques_ignores(fuseau):
    class Client:
        def free_busy(self, debut, pas):
            return "0?2"

    grille = SourceOutlook(Client(), fuseau).grille([date(2026, 11, 2)])
    assert len(grille.codes) == 2


class FausseComError(Exception):
    pass


def test_erreur_outlook_ferme():
    exc = FausseComError(MK_E_UNAVAILABLE, "Opération non disponible", None, None)
    assert isinstance(traduire_erreur_com(exc, "connexion"), ErreurOutlookFerme)


def test_erreur_alerte_securite_refusee():
    exc = FausseComError(DISP_E_EXCEPTION, "Exception", (0, "Outlook", "refus", None, 0, E_ABORT), None)
    erreur = traduire_erreur_com(exc, "lecture libre/occupé")
    assert isinstance(erreur, ErreurSecuriteOutlook)
    assert "Autoriser" in str(erreur)


def test_erreur_outlook_generique_affiche_le_code():
    exc = FausseComError(-2147467259, "Erreur non spécifiée", None, None)
    erreur = traduire_erreur_com(exc, "lecture libre/occupé")
    assert type(erreur) is ErreurOutlook
    assert "0x80004005" in str(erreur)
