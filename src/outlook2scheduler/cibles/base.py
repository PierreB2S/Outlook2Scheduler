"""Interface commune des « cibles sondage » (When2meet aujourd'hui, Framadate/Doodle plus tard)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Protocol


@dataclass(frozen=True)
class Sondage:
    url: str
    titre: str
    creneaux: list[int]  # epochs de début, dans l'ordre du sondage
    duree_creneau: int = 15 * 60  # secondes
    # Disponibilités déjà en ligne : nom du participant -> indices des créneaux où il est libre
    participants: dict[str, frozenset[int]] = field(default_factory=dict)

    @property
    def fin(self) -> int:
        return self.creneaux[-1] + self.duree_creneau


@dataclass(frozen=True)
class ResultatEnvoi:
    identifiant_participant: str
    nb_libres: int
    verifie: bool | None  # None = vérification impossible
    detail: str = ""


class CibleSondage(Protocol):
    nom: str

    def reconnait(self, url: str) -> bool: ...

    def normaliser(self, url: str) -> str:
        """Forme canonique du lien (pour reconnaître un sondage déjà envoyé)."""
        ...

    def lire(self, url: str) -> Sondage: ...

    def envoyer(
        self, sondage: Sondage, nom: str, mot_de_passe: str, disponibilites: list[bool]
    ) -> ResultatEnvoi: ...
