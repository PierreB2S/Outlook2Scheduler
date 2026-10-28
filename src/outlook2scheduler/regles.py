"""Règles de disponibilité : transforment la grille d'agenda en oui/non par créneau."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, time
from enum import Enum

from outlook2scheduler.erreurs import ErreurConfig
from outlook2scheduler.sources.base import (
    ABSENT,
    AILLEURS,
    LIBRE,
    OCCUPE,
    PAS_SECONDES,
    PROVISOIRE,
    GrilleAgenda,
)
from outlook2scheduler.temps import Fuseau

NOMS_JOURS = {
    "lun": 0, "lundi": 0,
    "mar": 1, "mardi": 1,
    "mer": 2, "mercredi": 2,
    "jeu": 3, "jeudi": 3,
    "ven": 4, "vendredi": 4,
    "sam": 5, "samedi": 5,
    "dim": 6, "dimanche": 6,
}


@dataclass(frozen=True)
class Regles:
    heure_debut: time = time(8, 30)
    heure_fin: time = time(18, 0)
    jours_exclus: frozenset[int] = field(default_factory=lambda: frozenset({5, 6}))
    provisoire_libre: bool = False
    ailleurs_libre: bool = True
    marge_minutes: int = 0

    def __post_init__(self) -> None:
        if self.heure_fin <= self.heure_debut and self.heure_fin != time(0):
            raise ErreurConfig("La plage horaire est vide : l'heure de fin doit suivre l'heure de début.")
        if self.marge_minutes < 0:
            raise ErreurConfig("La marge doit être positive ou nulle.")

    def decrire(self) -> str:
        jours = ", ".join(k for k, v in NOMS_JOURS.items() if len(k) == 3 and v in self.jours_exclus)
        morceaux = [
            f"{self.heure_debut:%H:%M}-{self.heure_fin:%H:%M}",
            f"jours exclus : {jours or 'aucun'}",
            f"provisoire = {'libre' if self.provisoire_libre else 'occupé'}",
        ]
        if self.marge_minutes:
            morceaux.append(f"marge {self.marge_minutes} min avant/après")
        return " | ".join(morceaux)


def parse_heures(texte: str) -> tuple[time, time]:
    m = re.fullmatch(r"\s*(\d{1,2})[:hH](\d{2})\s*-\s*(\d{1,2})[:hH](\d{2})\s*", texte or "")
    if not m:
        raise ErreurConfig(f"Plage horaire illisible : {texte!r} (attendu par ex. 08:30-18:00).")

    h1, m1, h2, m2 = (int(x) for x in m.groups())
    try:
        debut = time(h1, m1)
        fin = time(0) if (h2, m2) == (24, 0) else time(h2, m2)
    except ValueError as exc:
        raise ErreurConfig(f"Plage horaire invalide : {texte!r}.") from exc
    return debut, fin


def parse_jours(valeur: str | list[str] | None) -> frozenset[int]:
    if valeur is None:
        return frozenset()
    elements = valeur.split(",") if isinstance(valeur, str) else list(valeur)
    jours: set[int] = set()
    for brut in elements:
        cle = str(brut).strip().lower().rstrip(".")
        if not cle or cle == "aucun":
            continue
        if cle not in NOMS_JOURS:
            raise ErreurConfig(f"Jour inconnu : {brut!r} (utilisez lun, mar, mer, jeu, ven, sam, dim).")
        jours.add(NOMS_JOURS[cle])
    return frozenset(jours)


class Motif(str, Enum):
    LIBRE = "libre"
    OCCUPE = "occupé"
    MARGE = "marge autour d'une réunion"
    HORS_PLAGE = "hors plage horaire"
    JOUR_EXCLU = "jour exclu"
    HORS_GRILLE = "hors de la grille Outlook"


@dataclass(frozen=True)
class Creneau:
    epoch: int
    debut_local: datetime
    fin_local: datetime
    libre: bool
    motif: Motif
    changement_heure: bool = False

    @property
    def jour(self) -> date:
        return self.debut_local.date()


def _est_occupe(code: int, regles: Regles) -> bool:
    if code == LIBRE:
        return False
    if code == PROVISOIRE:
        return not regles.provisoire_libre
    if code == AILLEURS:
        return not regles.ailleurs_libre
    return code in (OCCUPE, ABSENT) or code > AILLEURS


def _minutes(t: time) -> int:
    return t.hour * 60 + t.minute


def _dans_plage(debut: datetime, duree_s: int, regles: Regles) -> bool:
    d = _minutes(debut.time())
    f = d + duree_s // 60
    fin_plage = _minutes(regles.heure_fin) or 24 * 60
    return d >= _minutes(regles.heure_debut) and f <= fin_plage


def _quarts(debut: int, fin: int) -> range:
    premier = debut - debut % PAS_SECONDES
    return range(premier, fin, PAS_SECONDES)


def evaluer(
    slots: list[int],
    grille: GrilleAgenda,
    regles: Regles,
    fuseau: Fuseau,
    duree_slot: int = PAS_SECONDES,
) -> list[Creneau]:
    marge = regles.marge_minutes * 60
    resultat: list[Creneau] = []
    for epoch in slots:
        debut = fuseau.vers_local(epoch)
        fin = fuseau.vers_local(epoch + duree_slot)
        dst = debut.date() in grille.jours_changement_heure

        def creneau(libre: bool, motif: Motif) -> Creneau:
            return Creneau(epoch, debut, fin, libre, motif, dst)

        if debut.weekday() in regles.jours_exclus:
            resultat.append(creneau(False, Motif.JOUR_EXCLU))
            continue
        if not _dans_plage(debut, duree_slot, regles):
            resultat.append(creneau(False, Motif.HORS_PLAGE))
            continue

        codes_slot = [grille.code(q) for q in _quarts(epoch, epoch + duree_slot)]
        if any(c is None for c in codes_slot):
            resultat.append(creneau(False, Motif.HORS_GRILLE))
            continue
        if any(_est_occupe(c, regles) for c in codes_slot):
            resultat.append(creneau(False, Motif.OCCUPE))
            continue
        if marge:
            autour = [
                grille.code(q)
                for q in _quarts(epoch - marge, epoch + duree_slot + marge)
                if not (epoch <= q < epoch + duree_slot)
            ]
            if any(c is not None and _est_occupe(c, regles) for c in autour):
                resultat.append(creneau(False, Motif.MARGE))
                continue
        resultat.append(creneau(True, Motif.LIBRE))
    return resultat
