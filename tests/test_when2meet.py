from datetime import date, time

import pytest
import requests
from conftest import FausseReponse, FausseSession, epochs_sondage, page_when2meet

from outlook2when2meet.cibles.when2meet import (
    When2meet,
    analyser_page,
    analyser_url,
    chaine_disponibilites,
)
from outlook2when2meet.erreurs import (
    ErreurFormatPage,
    ErreurMotDePasse,
    ErreurReseau,
    ErreurSondage,
    ErreurSondageJoursSemaine,
    ErreurSSL,
)

URL = "https://www.when2meet.com/?12345678-AbCdE"


@pytest.fixture
def epochs(fuseau):
    return epochs_sondage(fuseau, [date(2026, 11, 2), date(2026, 11, 3)], time(9), time(10))


@pytest.mark.parametrize(
    "url",
    [URL, "when2meet.com/?12345678-AbCdE", "  https://when2meet.com/?12345678-AbCdE#x", "http://www.when2meet.com/index.php?12345678-AbCdE"],
)
def test_analyser_url(url):
    assert analyser_url(url) == ("12345678", "AbCdE")


@pytest.mark.parametrize("url", ["", "https://framadate.org/abc", "https://www.when2meet.com/"])
def test_url_refusee(url):
    with pytest.raises(ErreurSondage):
        analyser_url(url)


def test_analyser_page(epochs):
    s = analyser_page(page_when2meet(epochs), URL)
    assert s.creneaux == epochs
    assert s.duree_creneau == 900
    assert s.titre == "Réunion test & démo"


def test_page_sans_creneaux():
    with pytest.raises(ErreurFormatPage):
        analyser_page("<html>Nouvelle maquette du site</html>", URL)


def test_pas_inattendu():
    with pytest.raises(ErreurFormatPage):
        analyser_page(page_when2meet([1_800_000_000, 1_800_001_800]), URL)


def test_sondage_jours_de_la_semaine_par_libelles(epochs):
    with pytest.raises(ErreurSondageJoursSemaine):
        analyser_page(page_when2meet(epochs, jours_semaine=True), URL)


def test_sondage_jours_de_la_semaine_par_dates_1978():
    fictives = [279_795_600 + 900 * i for i in range(8)]  # dates de 1978
    with pytest.raises(ErreurSondageJoursSemaine):
        analyser_page(page_when2meet(fictives), URL)


def test_chaine_disponibilites():
    assert chaine_disponibilites([True, False, True]) == "101"


def _cible(epochs, apres_envoi_libres, login="4242"):
    session = FausseSession()
    session.programmer(
        "/?12345678-AbCdE",
        page_when2meet(epochs, dispo={"999": [0]}),
        page_when2meet(epochs, dispo={"999": [0], "4242": apres_envoi_libres}),
    )
    session.programmer("/ProcessLogin.php", login)
    session.programmer("/SaveTimes.php", "")
    return When2meet(session), session


def test_envoi_complet_et_verifie(epochs):
    cible, session = _cible(epochs, [1, 2])
    sondage = cible.lire(URL)
    dispo = [False, True, True, False] + [False] * 4
    resultat = cible.envoyer(sondage, "Testeur", "", dispo)

    login = session.posts("/ProcessLogin.php")[0]
    assert login == {"id": "12345678", "name": "Testeur", "password": ""}
    save = session.posts("/SaveTimes.php")[0]
    assert save["person"] == "4242"
    assert save["event"] == "12345678"
    assert save["availability"] == "01100000"
    assert save["slots"] == f"{epochs[1]},{epochs[2]}"
    assert save["ChangeToAvailable"] == "true"
    assert resultat.verifie is True and resultat.nb_libres == 2


def test_envoi_non_persiste_detecte(epochs):
    cible, _ = _cible(epochs, [])
    sondage = cible.lire(URL)
    resultat = cible.envoyer(sondage, "Testeur", "", [True] + [False] * 7)
    assert resultat.verifie is False
    assert "1 manquants" in resultat.detail


def test_mot_de_passe_refuse(epochs):
    cible, session = _cible(epochs, [], login="Wrong Password.")
    sondage = cible.lire(URL)
    with pytest.raises(ErreurMotDePasse):
        cible.envoyer(sondage, "Testeur", "faux", [False] * 8)
    assert session.posts("/SaveTimes.php") == []


def test_reponse_de_connexion_inattendue(epochs):
    cible, _ = _cible(epochs, [], login="<html>maintenance</html>")
    sondage = cible.lire(URL)
    with pytest.raises(ErreurFormatPage):
        cible.envoyer(sondage, "Testeur", "", [False] * 8)


def test_erreur_ssl_du_proxy():
    session = FausseSession()
    session.programmer("/?12345678-AbCdE", requests.exceptions.SSLError("CERTIFICATE_VERIFY_FAILED"))
    with pytest.raises(ErreurSSL, match="proxy"):
        When2meet(session).lire(URL)


def test_erreur_reseau():
    session = FausseSession()
    session.programmer("/?12345678-AbCdE", requests.exceptions.ConnectionError("injoignable"))
    with pytest.raises(ErreurReseau):
        When2meet(session).lire(URL)


def test_erreur_http():
    session = FausseSession()
    session.programmer("/?12345678-AbCdE", FausseReponse(status_code=503))
    with pytest.raises(ErreurReseau, match="503"):
        When2meet(session).lire(URL)
