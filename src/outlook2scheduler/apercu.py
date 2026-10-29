"""Aperçu texte des plages libres et des changements, affiché avant tout envoi."""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from typing import Callable

from outlook2scheduler.cibles.base import Sondage
from outlook2scheduler.regles import Creneau, Motif, Regles
from outlook2scheduler.temps import JOURS_COURTS

Plages = dict[date, list[tuple[datetime, datetime]]]


def _jour(d: date) -> str:
    return f"{JOURS_COURTS[d.weekday()]} {d:%d/%m/%Y}"


def _plages(creneaux: list[Creneau], garder: Callable[[int, Creneau], bool]) -> Plages:
    """Regroupe en plages, jour par jour, les créneaux consécutifs retenus par ``garder``."""
    plages: Plages = {c.jour: [] for c in creneaux}
    precedent: Creneau | None = None
    precedent_garde = False
    for i, c in enumerate(creneaux):
        garde = garder(i, c)
        if garde:
            liste = plages[c.jour]
            if precedent_garde and precedent is not None and precedent.jour == c.jour \
                    and precedent.fin_local == c.debut_local and liste:
                liste[-1] = (liste[-1][0], c.fin_local)
            else:
                liste.append((c.debut_local, c.fin_local))
        precedent, precedent_garde = c, garde
    return dict(sorted(plages.items()))


def plages_libres(creneaux: list[Creneau]) -> Plages:
    return _plages(creneaux, lambda _i, c: c.libre)


def _texte_plages(plages: list[tuple[datetime, datetime]]) -> str:
    return ", ".join(f"{a:%H:%M}-{b:%H:%M}" for a, b in plages)


def formater(
    sondage: Sondage,
    creneaux: list[Creneau],
    regles: Regles,
    nb_rendez_vous: int | None = None,
    detail: bool = True,
    alertes: list[str] | None = None,
) -> str:
    lignes = [
        f"Sondage : {sondage.titre}",
        f"Règles  : {regles.decrire()}",
    ]
    if nb_rendez_vous is not None:
        lignes.append(f"Agenda  : {nb_rendez_vous} rendez-vous non libre(s) sur la période (horaires et statut uniquement)")
    if detail:
        lignes += ["", "Plages libres (heure de ce poste) :"]
        for jour, plages in plages_libres(creneaux).items():
            lignes.append(f"  {_jour(jour)} : {_texte_plages(plages) or '-- aucune --'}")

    nb_libres = sum(c.libre for c in creneaux)
    lignes += ["", f"Total : {nb_libres} créneau(x) de 15 min libres sur {len(creneaux)}."]

    motifs = Counter(c.motif for c in creneaux)
    avertissements: list[str] = list(alertes or [])
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
    if nb_rendez_vous == 0:
        avertissements.append(
            "Aucun rendez-vous lu dans Outlook sur la période. Si votre agenda n'est pas vide, "
            "répondez N et vérifiez avec : o2s.bat --agenda AAAA-MM-JJ"
        )
    if nb_libres == 0:
        avertissements.append("Aucun créneau libre : l'envoi effacerait vos disponibilités existantes.")
    if avertissements:
        lignes += ["", "Attention :"] + [f"  - {a}" for a in avertissements]
    return "\n".join(lignes)


def formater_ecart(creneaux: list[Creneau], avant: frozenset[int] | None, nom: str) -> str:
    """Compare avec ce qui est déjà en ligne sous ce nom."""
    if avant is None:
        return f"Première participation sous le nom « {nom} » dans ce sondage."
    apres = {i for i, c in enumerate(creneaux) if c.libre}
    if apres == set(avant):
        return f"Déjà à jour : ce qui est en ligne sous « {nom} » correspond à votre agenda."
    lignes = [f"Changements par rapport à ce qui est en ligne sous « {nom} » "
              f"({len(avant)} -> {len(apres)} créneaux libres) :"]
    for titre, garder in (
        ("Devenus OCCUPÉS", lambda i, c: i in avant and not c.libre),
        ("Devenus libres ", lambda i, c: i not in avant and c.libre),
    ):
        for jour, plages in _plages(creneaux, garder).items():
            if plages:
                lignes.append(f"  {titre} : {_jour(jour)} {_texte_plages(plages)}")
    return "\n".join(lignes)
