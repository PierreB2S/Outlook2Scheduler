"""Accès HTTP. Deux modes au choix :

- « windows » (par défaut sous Windows) : passe par WinINet, la pile réseau
  d'Edge et d'Excel (objet COM MSXML2.XMLHTTP). Elle suit les mêmes réglages
  proxy que le navigateur, y compris un script de configuration automatique
  (PAC) ou la détection automatique. Elle s'authentifie au proxy avec la
  session Windows et fait confiance aux certificats du magasin Windows.
- « python » : bibliothèque requests + truststore. Ne gère que les proxys
  fixes (variable HTTPS_PROXY ou proxy manuel déclaré dans Windows).

Les deux exposent ``request(methode, url, timeout=..., data=..., headers=...)``
et renvoient un objet avec ``status_code`` et ``text``.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Any, Callable
from urllib.parse import urlencode

from outlook2scheduler import __version__
from outlook2scheduler.erreurs import ErreurConfig, ErreurReseau, ErreurSSL

MODES = ("windows", "python")
URL_TEST = "https://www.when2meet.com/"

# Codes WinINet (HRESULT 0x8007xxxx) liés aux certificats
_ERREURS_CERTIFICAT = {12037, 12038, 12044, 12045, 12055, 12056, 12057, 12169, 12175, 12179}
_ERREURS_NOM = {12007}  # ERROR_INTERNET_NAME_NOT_RESOLVED
_ERREURS_DELAI = {12002}  # ERROR_INTERNET_TIMEOUT
_ERREURS_CONNEXION = {12029, 12030, 12031}


def mode_par_defaut() -> str:
    return "windows" if sys.platform == "win32" else "python"


@dataclass
class ReponseHTTP:
    status_code: int
    text: str


def _code_wininet(exc: BaseException) -> int | None:
    """Extrait le code WinINet (12xxx) d'une com_error, s'il y en a un."""
    args = getattr(exc, "args", ())
    candidats = []
    if args and isinstance(args[0], int):
        candidats.append(args[0])
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 5 and isinstance(args[2][5], int):
        candidats.append(args[2][5])
    for hr in candidats:
        hr &= 0xFFFFFFFF
        if hr >> 16 == 0x8007 and 12000 <= (hr & 0xFFFF) < 13000:
            return hr & 0xFFFF
    return None


def traduire_erreur_windows(exc: BaseException, url: str) -> ErreurReseau:
    code = _code_wininet(exc)
    detail = f"code WinINet {code}" if code else str(exc)
    if code in _ERREURS_CERTIFICAT:
        return ErreurSSL(
            f"Certificat refusé en contactant {url} ({detail}). Le certificat du proxy "
            "d'entreprise n'est pas reconnu par Windows : contactez l'IT."
        )
    if code in _ERREURS_NOM:
        return ErreurReseau(
            f"Nom de site introuvable en contactant {url} ({detail}). Vérifiez que le "
            "site s'ouvre dans Edge sur ce poste."
        )
    if code in _ERREURS_DELAI:
        return ErreurReseau(f"Délai dépassé en contactant {url} ({detail}).")
    if code in _ERREURS_CONNEXION:
        return ErreurReseau(f"Connexion interrompue avec {url} ({detail}). Réessayez.")
    return ErreurReseau(f"Impossible de joindre {url} ({detail}).")


class SessionWindows:
    """Requêtes HTTP via WinINet (MSXML2.XMLHTTP), comme Edge ou une macro Excel."""

    def __init__(self, fabrique_com: Callable[[], Any] | None = None) -> None:
        self._fabrique = fabrique_com or self._fabrique_reelle
        self._com_error: type[BaseException] = Exception

    def _fabrique_reelle(self) -> Any:  # pragma: no cover - dépend de Windows
        try:
            import pywintypes
            import win32com.client
        except ImportError as exc:
            raise ErreurConfig(
                "Le mode réseau « windows » nécessite pywin32 (Windows uniquement). "
                "Utilisez --reseau python."
            ) from exc
        self._com_error = pywintypes.com_error
        for progid in ("MSXML2.XMLHTTP.6.0", "MSXML2.XMLHTTP"):
            try:
                return win32com.client.Dispatch(progid)
            except pywintypes.com_error:
                continue
        raise ErreurReseau("Composant Windows MSXML2.XMLHTTP indisponible : utilisez --reseau python.")

    def request(
        self,
        methode: str,
        url: str,
        timeout: float | None = None,
        data: dict[str, str] | None = None,
        headers: dict[str, str] | None = None,
    ) -> ReponseHTTP:
        xhr = self._fabrique()
        try:
            xhr.open(methode, url, False)
            entetes = {
                # Évite que WinINet ne resserve une page en cache (relecture après envoi)
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
                "If-Modified-Since": "Sat, 01 Jan 2000 00:00:00 GMT",
                **(headers or {}),
            }
            corps = ""
            if data is not None:
                entetes["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
                corps = urlencode(data)
            for nom, valeur in entetes.items():
                try:
                    xhr.setRequestHeader(nom, valeur)
                except self._com_error:
                    pass  # certains en-têtes (Referer, Origin) peuvent être refusés : sans gravité
            xhr.send(corps)
            return ReponseHTTP(int(xhr.status), str(xhr.responseText or ""))
        except self._com_error as exc:
            raise traduire_erreur_windows(exc, url) from exc


def session_python() -> Any:  # pragma: no cover - dépend du réseau du poste
    import requests
    import truststore

    truststore.inject_into_ssl()
    session = requests.Session()
    session.headers["User-Agent"] = f"outlook2scheduler/{__version__}"
    return session


def creer_session(mode: str) -> Any:  # pragma: no cover
    if mode not in MODES:
        raise ErreurConfig(f"Mode réseau inconnu : {mode!r} (choix : {', '.join(MODES)}).")
    return SessionWindows() if mode == "windows" else session_python()


def verifier_acces(session: Any, url: str = URL_TEST) -> str:
    """Fait un simple GET et renvoie un compte rendu (lève ErreurReseau en cas d'échec)."""
    reponse = session.request("GET", url, timeout=30)
    if reponse.status_code == 407:
        raise ErreurReseau("Le proxy demande une authentification (HTTP 407).")
    if reponse.status_code >= 400:
        raise ErreurReseau(f"Réponse HTTP {reponse.status_code}.")
    return f"HTTP {reponse.status_code}, {len(reponse.text)} caractères reçus"
