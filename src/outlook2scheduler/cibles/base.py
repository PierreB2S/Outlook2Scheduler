"""Interface commune des « cibles sondage » (When2meet aujourd'hui, Framadate/Doodle plus tard)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class Sondage:
    url: str
    titre: str
    creneaux: list[int]  # epochs de début, dans l'ordre du sondage
    duree_creneau: int = 15 * 60  # secondes


@dataclass(frozen=True)
class ResultatEnvoi:
    identifiant_participant: str
    nb_libres: int
    verifie: bool | None  # None = vérification impossible
    detail: str = ""


class CibleSondage(Protocol):
    nom: str

    def reconnait(self, url: str) -> bool: ...

    def lire(self, url: str) -> Sondage: ...

    def envoyer(
        self, sondage: Sondage, nom: str, mot_de_passe: str, disponibilites: list[bool]
    ) -> ResultatEnvoi: ...
