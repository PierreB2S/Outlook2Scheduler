"""Lecture de config.toml (non versionné ; modèle : config.example.toml)."""

from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from outlook2scheduler.erreurs import ErreurConfig

FICHIER_PAR_DEFAUT = Path("config.toml")

CLES = {
    "nom": str,
    "mot_de_passe": str,
    "heures": (list, str),
    "jours_exclus": (list, str),
    "provisoire_libre": bool,
    "ailleurs_libre": bool,
    "marge_minutes": int,
    "reseau": str,
    "source_agenda": str,
}


@dataclass
class Config:
    nom: str = ""
    mot_de_passe: str = ""
    heures: list[str] | str = "08:30-18:00"  # une plage, ou plusieurs pour exclure la pause de midi
    jours_exclus: list[str] | str = ("sam", "dim")  # type: ignore[assignment]
    provisoire_libre: bool = False
    ailleurs_libre: bool = True
    marge_minutes: int = 0
    reseau: str = ""  # vide = mode par défaut (windows sous Windows)
    source_agenda: str = "rendez-vous"  # ou "freebusy"


def charger(chemin: Path | None = None) -> Config:
    fichier = chemin or FICHIER_PAR_DEFAUT
    if not fichier.exists():
        if chemin is not None:
            raise ErreurConfig(f"Fichier de configuration introuvable : {fichier}")
        return Config()
    try:
        donnees: dict[str, Any] = tomllib.loads(fichier.read_text(encoding="utf-8-sig"))  # utf-8-sig : tolère le BOM du Bloc-notes
    except (tomllib.TOMLDecodeError, UnicodeDecodeError) as exc:
        raise ErreurConfig(f"{fichier} est mal formé : {exc}") from exc

    config = Config()
    for cle, valeur in donnees.items():
        if cle not in CLES:
            raise ErreurConfig(
                f"Clé inconnue dans {fichier} : {cle!r} (clés possibles : {', '.join(CLES)})."
            )
        attendu = CLES[cle]
        if not isinstance(valeur, attendu) or (attendu is int and isinstance(valeur, bool)):
            raise ErreurConfig(f"Valeur invalide pour {cle!r} dans {fichier} : {valeur!r}.")
        setattr(config, cle, valeur)
    return config
