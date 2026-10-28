"""Aperçu texte des plages libres, affiché avant tout envoi."""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime

from outlook2when2meet.cibles.base import Sondage
from outlook2when2meet.regles import Creneau, Motif, Regles
from outlook2when2meet.temps import JOURS_COURTS


def _jour(d: date) -> str:
    return f"{JOURS_COURTS[d.weekday()]} {d:%d/%m/%Y}"


def plages_libres(creneaux: list[Creneau]) -> dict[date, list[tuple[datetime, datetime]]]:
    """Regroupe les créneaux libres consécutifs en plages, jour par jour."""
    plages: dict[date, list[tuple[datetime, datetime]]] = {c.jour: [] for c in creneaux}
    precedent: Creneau | None = None
    for c in creneaux:
        if c.libre:
            liste = plages[c.jour]
            if (
                precedent is not None
                and precedent.libre
                and precedent.jour == c.jour
                and precedent.fin_local == c.debut_local
                and liste
            ):
                liste[-1] = (liste[-1][0], c.fin_local)
            else:
                liste.append((c.debut_local, c.fin_local))
        precedent = c
    return dict(sorted(plages.items()))


def formater(sondage: Sondage, creneaux: list[Creneau], regles: Regles) -> str:
    lignes = [
        f"Sondage : {sondage.titre}",
        f"Règles  : {regles.decrire()}",
        "",
        "Plages libres (heure de ce poste) :",
    ]
    for jour, plages in plages_libres(creneaux).items():
        texte = ", ".join(f"{a:%H:%M}-{b:%H:%M}" for a, b in plages) or "-- aucune --"
        lignes.append(f"  {_jour(jour)} : {texte}")

    nb_libres = sum(c.libre for c in creneaux)
    lignes += ["", f"Total : {nb_libres} créneau(x) de 15 min libres sur {len(creneaux)}."]

    motifs = Counter(c.motif for c in creneaux)
    avertissements: list[str] = []
    if motifs[Motif.HORS_GRILLE]:
        avertissements.append(
            f"{motifs[Motif.HORS_GRILLE]} créneau(x) hors de la grille Outlook "
            "(trop loin dans le futur ou non publiés) : marqués occupés."
        )
    jours_dst = sorted({c.jour for c in creneaux if c.changement_heure})
    for jour in jours_dst:
        avertissements.append(
            f"{_jour(jour)} est un jour de changement d'heure : vérifiez ce jour-là à la main."
        )
    if nb_libres == 0:
        avertissements.append("Aucun créneau libre : l'envoi effacerait vos disponibilités existantes.")
    if avertissements:
        lignes += ["", "Attention :"] + [f"  - {a}" for a in avertissements]
    return "\n".join(lignes)
