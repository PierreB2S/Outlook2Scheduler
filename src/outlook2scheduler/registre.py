"""Mémoire locale des sondages envoyés, pour pouvoir les actualiser plus tard.

Fichier ``sondages.json`` à côté de config.toml, jamais versionné. Il ne
contient que : lien du sondage, titre, nom utilisé, fin du sondage, date du
dernier envoi. Aucune donnée d'agenda.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

from outlook2scheduler.erreurs import ErreurConfig

FICHIER_PAR_DEFAUT = Path("sondages.json")


@dataclass
class EntreeSondage:
    url: str
    titre: str
    nom: str
    fin: int  # epoch de fin du dernier créneau
    dernier_envoi: str  # date ISO locale


class Registre:
    def __init__(self, chemin: Path = FICHIER_PAR_DEFAUT) -> None:
        self.chemin = chemin

    def lire(self) -> list[EntreeSondage]:
        if not self.chemin.exists():
            return []
        try:
            donnees = json.loads(self.chemin.read_text(encoding="utf-8"))
            return [EntreeSondage(**e) for e in donnees]
        except (ValueError, TypeError) as exc:
            raise ErreurConfig(
                f"{self.chemin} est illisible ({exc}). Supprimez-le : seuls les "
                "sondages à actualiser seront oubliés."
            ) from exc

    def _ecrire(self, entrees: list[EntreeSondage]) -> None:
        self.chemin.write_text(
            json.dumps([asdict(e) for e in entrees], ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def enregistrer(self, url: str, titre: str, nom: str, fin: int, maintenant: float) -> None:
        entrees = [e for e in self.lire() if e.url != url]
        date_envoi = datetime.fromtimestamp(maintenant).isoformat(timespec="minutes")
        entrees.append(EntreeSondage(url, titre, nom, fin, date_envoi))
        self._ecrire(entrees)

    def oublier(self, url: str) -> EntreeSondage | None:
        entrees = self.lire()
        restantes = [e for e in entrees if e.url != url]
        if len(restantes) == len(entrees):
            return None
        self._ecrire(restantes)
        return next(e for e in entrees if e.url == url)

    def purger_termines(self, maintenant: float) -> list[EntreeSondage]:
        """Retire les sondages dont tous les créneaux sont passés ; renvoie ceux retirés."""
        entrees = self.lire()
        termines = [e for e in entrees if e.fin <= maintenant]
        if termines:
            self._ecrire([e for e in entrees if e.fin > maintenant])
        return termines
