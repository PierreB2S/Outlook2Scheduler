from datetime import date, time
from urllib.parse import parse_qs

import pytest
from conftest import epochs_sondage, page_when2meet

from outlook2scheduler.cibles.when2meet import When2meet
from outlook2scheduler.erreurs import ErreurReseau, ErreurSSL
from outlook2scheduler.reseau import SessionWindows, verifier_acces, traduire_erreur_windows


class FausseComError(Exception):
    pass


def hresult_wininet(code: int) -> int:
    """HRESULT signé tel que renvoyé par pywin32 pour une erreur WinINet."""
    return (0x80070000 | code) - (1 << 32)


class FauxXHR:
    """Imite MSXML2.XMLHTTP : open / setRequestHeader / send / status / responseText."""

    def __init__(self, status=200, texte="", erreur=None, entetes_refuses=()):
        self.status, self.responseText = status, texte
        self.erreur, self.entetes_refuses = erreur, set(entetes_refuses)
        self.ouvert = self.corps = None
        self.entetes = {}

    def open(self, methode, url, asynchrone):
        assert asynchrone is False
        self.ouvert = (methode, url)

    def setRequestHeader(self, nom, valeur):
        if nom in self.entetes_refuses:
            raise FausseComError(-2147024809, "Paramètre incorrect", None, None)
        self.entetes[nom] = valeur

    def send(self, corps):
        if self.erreur:
            raise self.erreur
        self.corps = corps


def session_avec(*xhrs):
    file = list(xhrs)
    session = SessionWindows(fabrique_com=lambda: file.pop(0))
    session._com_error = FausseComError
    return session


def test_get_sans_cache():
    xhr = FauxXHR(texte="<html>ok</html>")
    r = session_avec(xhr).request("GET", "https://www.when2meet.com/?1-a", timeout=30)
    assert (r.status_code, r.text) == (200, "<html>ok</html>")
    assert xhr.ouvert == ("GET", "https://www.when2meet.com/?1-a")
    assert xhr.entetes["Cache-Control"] == "no-cache"
    assert "If-Modified-Since" in xhr.entetes
    assert xhr.corps == ""


def test_post_formulaire_encode():
    xhr = FauxXHR(texte="4242")
    data = {"name": "Prénom N.", "slots": "1,2", "password": ""}
    session_avec(xhr).request("POST", "https://x/ProcessLogin.php", data=data, headers={"Referer": "https://x/"})
    assert xhr.entetes["Content-Type"].startswith("application/x-www-form-urlencoded")
    assert xhr.entetes["Referer"] == "https://x/"
    assert parse_qs(xhr.corps, keep_blank_values=True) == {"name": ["Prénom N."], "slots": ["1,2"], "password": [""]}


def test_entete_refuse_sans_gravite():
    xhr = FauxXHR(entetes_refuses={"Referer"})
    r = session_avec(xhr).request("GET", "https://x/", headers={"Referer": "https://x/"})
    assert r.status_code == 200 and "Referer" not in xhr.entetes


@pytest.mark.parametrize(
    "code,classe,mot",
    [(12007, ErreurReseau, "introuvable"), (12175, ErreurSSL, "IT"), (12002, ErreurReseau, "Délai")],
)
def test_erreurs_wininet(code, classe, mot):
    exc = FausseComError(-2147352567, "Exception", (0, "msxml6.dll", "", None, 0, hresult_wininet(code)), None)
    erreur = traduire_erreur_windows(exc, "https://x/")
    assert type(erreur) is classe and mot in str(erreur) and str(code) in str(erreur)


def test_erreur_com_pendant_l_envoi():
    xhr = FauxXHR(erreur=FausseComError(hresult_wininet(12007), "", None, None))
    with pytest.raises(ErreurReseau, match="12007"):
        session_avec(xhr).request("GET", "https://x/")


def test_tester_signale_proxy_authentifiant():
    with pytest.raises(ErreurReseau, match="407"):
        verifier_acces(session_avec(FauxXHR(status=407)))
    assert "HTTP 200" in verifier_acces(session_avec(FauxXHR(texte="abc")))


def test_when2meet_de_bout_en_bout_via_windows(fuseau):
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    session = session_avec(
        FauxXHR(texte=page_when2meet(epochs)),
        FauxXHR(texte="4242"),
        sauve := FauxXHR(texte=""),
        FauxXHR(texte=page_when2meet(epochs, dispo={"4242": [0, 3]})),
    )
    cible = When2meet(session)
    sondage = cible.lire("https://www.when2meet.com/?12345678-AbCdE")
    resultat = cible.envoyer(sondage, "Testeur", "", [True, False, False, True])
    assert parse_qs(sauve.corps)["availability"] == ["1001"]
    assert resultat.verifie is True


def test_when2meet_proxy_407_via_python():
    from conftest import FausseReponse, FausseSession

    session = FausseSession()
    session.programmer("/?12345678-AbCdE", FausseReponse(status_code=407))
    with pytest.raises(ErreurReseau, match="authentification"):
        When2meet(session).lire("https://www.when2meet.com/?12345678-AbCdE")
