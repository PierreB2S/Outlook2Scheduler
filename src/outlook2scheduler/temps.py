"""Conversions entre instants (epoch UTC) et heure locale « murale » du poste.

Les sondages donnent des epochs, Outlook raisonne en heure locale : tout le
passage de l'un à l'autre passe par un ``Fuseau``, ce qui permet de tester
le changement d'heure avec un fuseau fixe (Europe/Paris) dans les tests.
"""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Protocol
from zoneinfo import ZoneInfo

_EPOCH_NAIF = datetime(1970, 1, 1)


class Fuseau(Protocol):
    def vers_local(self, epoch: int) -> datetime:
        """Epoch -> datetime locale naïve."""

    def vers_epoch(self, local: datetime) -> int:
        """Datetime locale naïve -> epoch."""


class FuseauSysteme:
    """Fuseau du poste Windows (règles d'heure d'été comprises)."""

    def vers_local(self, epoch: int) -> datetime:
        return datetime.fromtimestamp(epoch)

    def vers_epoch(self, local: datetime) -> int:
        return int(local.timestamp())


class FuseauZone:
    """Fuseau IANA explicite (utilisé par les tests)."""

    def __init__(self, nom: str) -> None:
        self.zone = ZoneInfo(nom)

    def vers_local(self, epoch: int) -> datetime:
        return datetime.fromtimestamp(epoch, self.zone).replace(tzinfo=None)

    def vers_epoch(self, local: datetime) -> int:
        return int(local.replace(tzinfo=self.zone).timestamp())


def minuit(jour: date) -> datetime:
    return datetime.combine(jour, time())


def decalage_utc(fuseau: Fuseau, local: datetime) -> int:
    """Décalage (en secondes) entre l'heure locale et UTC à cet instant."""
    return int((local - _EPOCH_NAIF).total_seconds()) - fuseau.vers_epoch(local)


def est_jour_changement_heure(fuseau: Fuseau, jour: date) -> bool:
    """Vrai si le décalage UTC change pendant ce jour (passage été/hiver)."""
    lendemain = jour + timedelta(days=1)
    return decalage_utc(fuseau, minuit(jour)) != decalage_utc(fuseau, minuit(lendemain))


JOURS_COURTS = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."]
