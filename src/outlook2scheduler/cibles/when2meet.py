"""Cible When2meet, via les endpoints HTTP qu'utilise la page elle-même.

- GET  /?<id>-<code>      : page du sondage, créneaux dans ``TimeOfSlot[i] = <epoch>``
- POST /ProcessLogin.php  : id, name, password -> identifiant numérique du participant
- POST /SaveTimes.php     : person, event, slots, availability (chaîne 0/1 complète),
                            password, ChangeToAvailable=true

Ces endpoints ne sont pas documentés : après l'envoi, on relit la page pour
vérifier que les disponibilités ont bien été enregistrées.
"""

from __future__ import annotations

import html as html_lib
import re
from typing import Any

from outlook2scheduler.cibles.base import ResultatEnvoi, Sondage
from outlook2scheduler.erreurs import (
    ErreurFormatPage,
    ErreurMotDePasse,
    ErreurReseau,
    ErreurSondage,
    ErreurSondageJoursSemaine,
    ErreurSSL,
)

BASE = "https://www.when2meet.com"
URL_RE = re.compile(r"when2meet\.com/(?:index\.php)?\?(\d+)-([A-Za-z0-9]+)", re.I)
TIME_OF_SLOT_RE = re.compile(r"TimeOfSlot\[(\d+)\]\s*=\s*(\d+)")
AVAILABLE_RE = re.compile(r"AvailableAtSlot\[(\d+)\]\.push\((\d+)\)")
TITRE_RE = re.compile(r"<title>\s*(.*?)\s*-\s*When2meet\s*</title>", re.S | re.I)
JOUR_SEMAINE_RE = re.compile(
    r'ShowSlot\(\d+,\s*"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b'
)
AN_2000 = 946684800  # les sondages « jours de la semaine » utilisent des dates de 1978

TIMEOUT = 30


def analyser_url(url: str) -> tuple[str, str]:
    m = URL_RE.search(url or "")
    if not m:
        raise ErreurSondage(
            f"Ce lien n'est pas un sondage When2meet : {url!r}\n"
            "Format attendu : https://www.when2meet.com/?12345678-AbCdE"
        )
    return m.group(1), m.group(2)


def url_canonique(url: str) -> str:
    ident, code = analyser_url(url)
    return f"{BASE}/?{ident}-{code}"


def analyser_page(page: str, url: str) -> Sondage:
    paires = sorted((int(i), int(t)) for i, t in TIME_OF_SLOT_RE.findall(page))
    if not paires:
        raise ErreurFormatPage(
            "Aucun créneau trouvé dans la page When2meet. Soit le lien est "
            "faux (sondage supprimé ?), soit When2meet a changé le format de "
            "sa page et l'outil doit être mis à jour."
        )
    if [i for i, _ in paires] != list(range(len(paires))):
        raise ErreurFormatPage("Numérotation des créneaux inattendue : When2meet a peut-être changé le format de sa page.")
    creneaux = [t for _, t in paires]
    if JOUR_SEMAINE_RE.search(page) or all(t < AN_2000 for t in creneaux):
        raise ErreurSondageJoursSemaine(
            "Ce sondage porte sur des « jours de la semaine » (lundi, mardi...) "
            "et non sur des dates précises : impossible de le relier à votre "
            "agenda. Remplissez-le à la main ou demandez un sondage par dates."
        )
    ecarts = [b - a for a, b in zip(creneaux, creneaux[1:]) if b > a]
    duree = min(ecarts) if ecarts else 900
    if duree != 900 or any(b <= a for a, b in zip(creneaux, creneaux[1:])):
        raise ErreurFormatPage(
            "Créneaux When2meet inattendus (pas de 15 min ou ordre incohérent) : "
            "le format de la page a peut-être changé."
        )
    m = TITRE_RE.search(page)
    titre = html_lib.unescape(m.group(1)) if m else "(sans titre)"
    return Sondage(url=url, titre=titre, creneaux=creneaux, duree_creneau=duree)


def chaine_disponibilites(disponibilites: list[bool]) -> str:
    return "".join("1" if d else "0" for d in disponibilites)


def _requete(session: Any, methode: str, url: str, **kwargs: Any) -> Any:
    import requests

    try:
        reponse = session.request(methode, url, timeout=TIMEOUT, **kwargs)
    except requests.exceptions.SSLError as exc:
        raise ErreurSSL(
            "Erreur de certificat SSL en contactant When2meet. C'est en général "
            "le proxy d'entreprise qui inspecte le trafic : l'outil utilise déjà "
            "les certificats Windows (truststore). Si l'erreur persiste, le "
            "certificat du proxy n'est pas dans le magasin Windows : contactez l'IT.\n"
            f"Détail : {exc}"
        ) from exc
    except requests.exceptions.ProxyError as exc:
        raise ErreurReseau(f"Le proxy a refusé la connexion à When2meet.\nDétail : {exc}") from exc
    except requests.exceptions.RequestException as exc:
        raise ErreurReseau(
            f"Impossible de joindre When2meet (réseau, proxy ou site indisponible).\nDétail : {exc}"
        ) from exc
    if reponse.status_code >= 400:
        raise ErreurReseau(f"When2meet a répondu HTTP {reponse.status_code} sur {url}.")
    return reponse


class When2meet:
    nom = "When2meet"

    def __init__(self, session: Any) -> None:
        self.session = session

    def reconnait(self, url: str) -> bool:
        return bool(URL_RE.search(url or ""))

    def _page(self, url: str) -> str:
        return _requete(self.session, "GET", url_canonique(url)).text

    def lire(self, url: str) -> Sondage:
        return analyser_page(self._page(url), url_canonique(url))

    def _connexion(self, ident: str, nom: str, mot_de_passe: str, referer: str) -> str:
        reponse = _requete(
            self.session,
            "POST",
            f"{BASE}/ProcessLogin.php",
            data={"id": ident, "name": nom, "password": mot_de_passe},
            headers={"Referer": referer, "Origin": BASE, "X-Requested-With": "XMLHttpRequest"},
        )
        texte = reponse.text.strip()
        if texte.isdigit():
            return texte
        if "password" in texte.lower():
            raise ErreurMotDePasse(
                f"When2meet a refusé le mot de passe pour « {nom} ». Ce nom existe "
                "déjà dans le sondage avec un autre mot de passe : indiquez le bon "
                "(--mot-de-passe ou config.toml) ou utilisez un autre nom."
            )
        raise ErreurFormatPage(
            "Réponse inattendue de When2meet à la connexion "
            f"({texte[:80]!r}) : le site a peut-être changé."
        )

    def envoyer(
        self, sondage: Sondage, nom: str, mot_de_passe: str, disponibilites: list[bool]
    ) -> ResultatEnvoi:
        if len(disponibilites) != len(sondage.creneaux):
            raise ValueError("une disponibilité par créneau est attendue")
        if not nom.strip():
            raise ErreurSondage("Le nom du participant est vide.")
        ident, _ = analyser_url(sondage.url)
        referer = url_canonique(sondage.url)
        personne = self._connexion(ident, nom.strip(), mot_de_passe, referer)
        libres = [t for t, d in zip(sondage.creneaux, disponibilites) if d]
        _requete(
            self.session,
            "POST",
            f"{BASE}/SaveTimes.php",
            data={
                "person": personne,
                "event": ident,
                "slots": ",".join(str(t) for t in libres),
                "availability": chaine_disponibilites(disponibilites),
                "password": mot_de_passe,
                "ChangeToAvailable": "true",
            },
            headers={"Referer": referer, "Origin": BASE, "X-Requested-With": "XMLHttpRequest"},
        )
        return self._verifier(sondage, personne, disponibilites)

    def _verifier(self, sondage: Sondage, personne: str, disponibilites: list[bool]) -> ResultatEnvoi:
        nb_libres = sum(disponibilites)
        try:
            page = self._page(sondage.url)
        except ErreurReseau as exc:
            return ResultatEnvoi(personne, nb_libres, None, f"relecture impossible : {exc}")
        enregistres = {int(i) for i, pid in AVAILABLE_RE.findall(page) if pid == personne}
        attendus = {i for i, d in enumerate(disponibilites) if d}
        if enregistres == attendus:
            return ResultatEnvoi(personne, nb_libres, True)
        return ResultatEnvoi(
            personne,
            nb_libres,
            False,
            f"{len(attendus)} créneaux envoyés, {len(enregistres)} lus sur la page "
            f"({len(attendus - enregistres)} manquants, {len(enregistres - attendus)} en trop).",
        )
