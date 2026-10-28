from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Iterable, Protocol

# Codes de disponibilité (ceux d'Outlook FreeBusy)
LIBRE = 0
PROVISOIRE = 1
OCCUPE = 2
ABSENT = 3
AILLEURS = 4

PAS_SECONDES = 15 * 60


@dataclass
class GrilleAgenda:
    """Disponibilité par quart d'heure, indexée par l'epoch de début du quart d'heure.

    Un quart d'heure absent de ``codes`` est « hors grille » (inconnu).
    """

    codes: dict[int, int] = field(default_factory=dict)
    jours_changement_heure: set[date] = field(default_factory=set)

    def code(self, epoch: int) -> int | None:
        return self.codes.get(epoch)


class SourceAgenda(Protocol):
    def grille(self, jours: Iterable[date]) -> GrilleAgenda:
        """Renvoie la grille couvrant au moins les jours (locaux) demandés."""
