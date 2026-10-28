from datetime import datetime, time

import pytest

from outlook2when2meet.erreurs import ErreurConfig
from outlook2when2meet.regles import Motif, Regles, evaluer, parse_heures, parse_jours
from outlook2when2meet.sources.base import ABSENT, AILLEURS, LIBRE, OCCUPE, PROVISOIRE, GrilleAgenda

# Lundi 2 novembre 2026
LUNDI = (2026, 11, 2)


def grille_journee(fuseau, jour=LUNDI, codes=None):
    """Grille complète d'une journée, libre sauf les quarts d'heure précisés {(h, m): code}."""
    g = GrilleAgenda()
    for q in range(96):
        h, m = divmod(q * 15, 60)
        g.codes[fuseau.vers_epoch(datetime(*jour, h, m))] = (codes or {}).get((h, m), LIBRE)
    return g


def slot(fuseau, h, m=0, jour=LUNDI):
    return fuseau.vers_epoch(datetime(*jour, h, m))


def un(fuseau, grille, h, m=0, jour=LUNDI, **regles):
    return evaluer([slot(fuseau, h, m, jour)], grille, Regles(**regles), fuseau)[0]


def test_libre(fuseau):
    c = un(fuseau, grille_journee(fuseau), 10)
    assert c.libre and c.motif is Motif.LIBRE
    assert c.debut_local == datetime(*LUNDI, 10) and c.fin_local == datetime(*LUNDI, 10, 15)


def test_plage_horaire(fuseau):
    g = grille_journee(fuseau)
    assert un(fuseau, g, 8, 15).motif is Motif.HORS_PLAGE
    assert un(fuseau, g, 8, 30).libre
    assert un(fuseau, g, 17, 45).libre  # se termine à 18:00 pile
    assert un(fuseau, g, 18, 0).motif is Motif.HORS_PLAGE


def test_plage_jusqu_a_minuit(fuseau):
    g = grille_journee(fuseau)
    assert un(fuseau, g, 23, 45, heure_debut=time(20), heure_fin=time(0)).libre


def test_jours_exclus(fuseau):
    samedi = (2026, 11, 7)
    g = grille_journee(fuseau, samedi)
    assert un(fuseau, g, 10, jour=samedi).motif is Motif.JOUR_EXCLU
    assert un(fuseau, g, 10, jour=samedi, jours_exclus=frozenset()).libre


@pytest.mark.parametrize(
    "code,attendu,options",
    [
        (OCCUPE, False, {}),
        (ABSENT, False, {}),
        (PROVISOIRE, False, {}),
        (PROVISOIRE, True, {"provisoire_libre": True}),
        (AILLEURS, True, {}),
        (AILLEURS, False, {"ailleurs_libre": False}),
    ],
)
def test_codes_outlook(fuseau, code, attendu, options):
    g = grille_journee(fuseau, codes={(10, 0): code})
    assert un(fuseau, g, 10, **options).libre is attendu


def test_marge_avant_apres(fuseau):
    g = grille_journee(fuseau, codes={(10, 0): OCCUPE, (10, 15): OCCUPE})  # réunion 10:00-10:30
    assert un(fuseau, g, 9, 45).libre
    assert un(fuseau, g, 9, 45, marge_minutes=15).motif is Motif.MARGE
    assert un(fuseau, g, 9, 30, marge_minutes=15).libre
    assert un(fuseau, g, 10, 30, marge_minutes=15).motif is Motif.MARGE
    assert un(fuseau, g, 10, 45, marge_minutes=15).libre
    assert un(fuseau, g, 10, 0, marge_minutes=15).motif is Motif.OCCUPE


def test_hors_grille_marque_occupe(fuseau):
    c = un(fuseau, GrilleAgenda(), 10)
    assert not c.libre and c.motif is Motif.HORS_GRILLE


def test_parse_heures():
    assert parse_heures("08:30-18:00") == (time(8, 30), time(18))
    assert parse_heures(" 9h00 - 17h30 ") == (time(9), time(17, 30))
    assert parse_heures("20:00-24:00") == (time(20), time(0))
    for faux in ["8-18", "25:00-26:00", ""]:
        with pytest.raises(ErreurConfig):
            parse_heures(faux)


def test_plage_vide_refusee():
    with pytest.raises(ErreurConfig):
        Regles(heure_debut=time(18), heure_fin=time(9))


def test_parse_jours():
    assert parse_jours("sam,dim") == frozenset({5, 6})
    assert parse_jours(["Lundi", "mer."]) == frozenset({0, 2})
    assert parse_jours("aucun") == frozenset()
    with pytest.raises(ErreurConfig):
        parse_jours("samedimanche")
