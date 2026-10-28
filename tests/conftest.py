"""Doublures de test : faux Outlook, fausses pages When2meet, fausse session HTTP.

Toutes les données sont synthétiques.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from typing import Any

import pytest

from outlook2scheduler.temps import FuseauZone

PARIS = "Europe/Paris"


@pytest.fixture
def fuseau() -> FuseauZone:
    return FuseauZone(PARIS)


# --------------------------------------------------------------------- Outlook
@dataclass
class FauxFreeBusy:
    """Imite Recipient.FreeBusy(debut, 15, True).

    ``evenements`` : liste de (debut_local, fin_local, code).
    ``semantique`` : « ecoule » (chaque caractère = 15 min de temps réel) ou
    « murale » (chaque caractère = 15 min d'horloge locale) ; on ne sait pas
    laquelle Outlook applique le jour du changement d'heure, le code doit
    être juste dans les deux cas hors de ce jour.
    ``horizon_jours`` : longueur de la chaîne renvoyée.
    """

    fuseau: FuseauZone
    evenements: list[tuple[datetime, datetime, int]] = field(default_factory=list)
    semantique: str = "ecoule"
    horizon_jours: int = 30
    appels: list[datetime] = field(default_factory=list)

    def _code(self, local: datetime) -> str:
        code = 0
        for debut, fin, c in self.evenements:
            if debut <= local < fin:
                code = max(code, c)
        return str(code)

    def free_busy(self, debut: datetime, pas_minutes: int) -> str:
        assert pas_minutes == 15
        self.appels.append(debut)
        depart = datetime.combine(debut.date(), time())
        n = self.horizon_jours * 96
        if self.semantique == "ecoule":
            base = self.fuseau.vers_epoch(depart)
            return "".join(self._code(self.fuseau.vers_local(base + i * 900)) for i in range(n))
        return "".join(self._code(depart + timedelta(minutes=15 * i)) for i in range(n))


# -------------------------------------------------------------------- When2meet
def epochs_sondage(fuseau: FuseauZone, jours: list[date], de: time, a: time) -> list[int]:
    """Epochs des créneaux de 15 min, de ``de`` à ``a`` (heure de Paris) pour chaque jour."""
    resultat = []
    for jour in jours:
        t = datetime.combine(jour, de)
        while t < datetime.combine(jour, a):
            resultat.append(fuseau.vers_epoch(t))
            t += timedelta(minutes=15)
    return resultat


def page_when2meet(
    epochs: list[int],
    titre: str = "Réunion test &amp; démo",
    dispo: dict[str, list[int]] | None = None,
    jours_semaine: bool = False,
    noms: dict[str, str] | None = None,
) -> str:
    """Page When2meet synthétique, reprenant la structure des variables JS de la vraie page.

    ``noms`` : identifiant -> nom affiché (déjà échappé comme en JavaScript).
    """
    lignes = [
        "<html><head>",
        f"<title>{titre} - When2meet</title>",
        "</head><body>",
        "<script>",
        "var TimeOfSlot=new Array(); var AvailableAtSlot=new Array();",
        "var PeopleNames = new Array(); var PeopleIDs = new Array();",
    ]
    for pid_index, pid in enumerate((dispo or {}).keys()):
        nom = (noms or {}).get(pid, f"Participant {pid_index}")
        lignes.append(f"PeopleNames[{pid_index}] = '{nom}';PeopleIDs[{pid_index}] = {pid};")
    for i, t in enumerate(epochs):
        lignes.append(f"TimeOfSlot[{i}]={t};")
        lignes.append(f"AvailableAtSlot[{i}]=new Array();")
    for pid, indices in (dispo or {}).items():
        for i in indices:
            lignes.append(f"AvailableAtSlot[{i}].push({pid});")
    lignes.append("</script>")
    if jours_semaine:
        lignes.append(f'<div onmouseover=\'ShowSlot({epochs[0]},"Monday 09:00:00 AM");\'></div>')
    lignes.append("</body></html>")
    return "\n".join(lignes)


@dataclass
class FausseReponse:
    text: str = ""
    status_code: int = 200


@dataclass
class FausseSession:
    """Remplace requests.Session : renvoie les réponses programmées, dans l'ordre, par URL."""

    reponses: dict[str, list[Any]] = field(default_factory=dict)
    requetes: list[tuple[str, str, dict[str, Any]]] = field(default_factory=list)

    def programmer(self, fin_url: str, *reponses: Any) -> None:
        self.reponses.setdefault(fin_url, []).extend(reponses)

    def request(self, methode: str, url: str, **kwargs: Any) -> FausseReponse:
        self.requetes.append((methode, url, kwargs))
        for fin_url, file in self.reponses.items():
            if url.endswith(fin_url) and file:
                r = file.pop(0) if len(file) > 1 else file[0]
                if isinstance(r, BaseException):
                    raise r
                return r if isinstance(r, FausseReponse) else FausseReponse(text=r)
        raise AssertionError(f"requête non prévue : {methode} {url}")

    def posts(self, fin_url: str) -> list[dict[str, Any]]:
        return [k["data"] for m, u, k in self.requetes if m == "POST" and u.endswith(fin_url)]
