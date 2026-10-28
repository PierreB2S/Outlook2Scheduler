"""Relances : écart avec ce qui est en ligne, mémoire des sondages, --actualiser / --liste / --oublier."""

import json
from datetime import date, datetime, time

import pytest
from conftest import FausseSession, epochs_sondage, page_when2meet

from outlook2scheduler.cibles.when2meet import When2meet, analyser_page
from outlook2scheduler.cli import Dependances, executer
from outlook2scheduler.registre import Registre
from outlook2scheduler.sources.base import OCCUPE
from outlook2scheduler.sources.outlook_rdv import RendezVous, SourceOutlookRendezVous

URL1 = "https://www.when2meet.com/?11111111-AaAaA"
URL2 = "https://www.when2meet.com/?22222222-BbBbB"
NOM = "Pierre T"
LUNDI, MARDI = date(2026, 11, 2), date(2026, 11, 3)


class Monde:
    def __init__(self, fuseau, pages, rdvs=(), reponses=(), maintenant=datetime(2026, 10, 1)):
        self.session = FausseSession()
        for suffixe, contenu in pages.items():
            self.session.programmer(suffixe, *(contenu if isinstance(contenu, tuple) else (contenu,)))
        self.session.programmer("/ProcessLogin.php", "4242")
        self.session.programmer("/SaveTimes.php", "")
        self.sorties, self.questions = [], []
        self._reponses = list(reponses)
        rdv_epochs = [RendezVous(fuseau.vers_epoch(d), fuseau.vers_epoch(f), OCCUPE) for d, f in rdvs]

        class Client:
            def rendez_vous(self, debut, fin):
                return rdv_epochs

        self.dep = Dependances(
            fabrique_source=lambda f, source: SourceOutlookRendezVous(Client(), f),
            fabrique_cibles=lambda mode: [When2meet(self.session)],
            fuseau=fuseau,
            entree=self._entree,
            sortie=self.sorties.append,
            presse_papiers=lambda: "",
            maintenant=lambda: fuseau.vers_epoch(maintenant),
        )

    def _entree(self, question):
        self.questions.append(question)
        return self._reponses.pop(0)

    @property
    def texte(self):
        return "\n".join(self.sorties)

    def envois(self):
        return [(d["event"], d["availability"]) for d in self.session.posts("/SaveTimes.php")]


def page(fuseau, jour, libres_en_ligne=None, titre="Point projet"):
    epochs = epochs_sondage(fuseau, [jour], time(9), time(10))
    if libres_en_ligne is None:
        return page_when2meet(epochs, titre=titre)
    return page_when2meet(epochs, titre=titre, dispo={"4242": libres_en_ligne}, noms={"4242": NOM})


@pytest.fixture(autouse=True)
def dossier_temporaire(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    return tmp_path


def test_ecart_affiche_puis_envoi(fuseau):
    m = Monde(fuseau, {"/?11111111-AaAaA": (page(fuseau, LUNDI, [0, 1, 2, 3]), page(fuseau, LUNDI, [0, 2, 3]))},
              rdvs=[(datetime(2026, 11, 2, 9, 15), datetime(2026, 11, 2, 9, 30))], reponses=["O"])
    executer([URL1, "--nom", NOM, "--submit"], m.dep)
    assert "(4 -> 3 créneaux libres)" in m.texte
    assert "Devenus OCCUPÉS : lun. 02/11/2026 09:15-09:30" in m.texte
    assert m.envois() == [("11111111", "1011")]
    assert "rechargez la page du sondage (F5)" in m.texte


def test_relance_o2s_ecrase_meme_si_deja_a_jour(fuseau):
    m = Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI, [0, 1, 2, 3])}, reponses=["O"])
    executer([URL1, "--nom", NOM, "--submit"], m.dep)
    assert "Déjà à jour" in m.texte
    assert "seront entièrement remplacées" in m.texte
    assert m.questions == ["Envoyer ? (O/N) "]
    assert m.envois() == [("11111111", "1111")]


def test_actualiser_deja_a_jour_rien_n_est_envoye(fuseau):
    Registre().enregistrer(URL1, "Point projet", NOM, 2_000_000_000, 0)
    m = Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI, [0, 1, 2, 3])})
    assert executer(["--actualiser", "--submit"], m.dep) == 0
    assert "Déjà à jour" in m.texte and "Rien à envoyer" in m.texte
    assert m.questions == [] and m.envois() == []
    assert [e.url for e in Registre().lire()] == [URL1]


def test_relance_reprend_le_nom_memorise(fuseau):
    # Premier envoi sous un nom saisi ; la relance (sans --nom ni config) réutilise ce nom.
    Registre().enregistrer(URL1, "Point projet", "Nom Saisi", 2_000_000_000, 0)
    m = Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI)}, reponses=["O"])
    executer([URL1, "--submit"], m.dep)
    assert m.session.posts("/ProcessLogin.php")[0]["name"] == "Nom Saisi"
    assert m.questions == ["Envoyer ? (O/N) "]


def test_premiere_participation(fuseau):
    m = Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI)}, reponses=["O"])
    executer([URL1, "--nom", NOM, "--submit"], m.dep)
    assert "Première participation" in m.texte
    assert m.envois() == [("11111111", "1111")]


def _monde_9h30(fuseau):
    # Il est 9h30 : 9h00 et 9h15 sont passés. En ligne : libre à 9h00 et 9h15 seulement.
    # Agenda : occupé de 9h00 à 9h30.
    return Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI, [0, 1])},
                 rdvs=[(datetime(2026, 11, 2, 9, 0), datetime(2026, 11, 2, 9, 30))],
                 reponses=["O"], maintenant=datetime(2026, 11, 2, 9, 30))


def test_actualiser_laisse_les_creneaux_passes_tels_qu_en_ligne(fuseau):
    Registre().enregistrer(URL1, "Point projet", NOM, 2_000_000_000, 0)
    m = _monde_9h30(fuseau)
    executer(["--actualiser", "--submit"], m.dep)
    assert m.envois() == [("11111111", "1111")]
    assert "Devenus libres  : lun. 02/11/2026 09:30-10:00" in m.texte
    assert "OCCUPÉS" not in m.texte


def test_relance_o2s_reecrit_tout_y_compris_le_passe(fuseau):
    m = _monde_9h30(fuseau)
    executer([URL1, "--nom", NOM, "--submit"], m.dep)
    assert m.envois() == [("11111111", "0011")]


def test_liste_et_oublier(fuseau):
    m = Monde(fuseau, {"/?11111111-AaAaA": page(fuseau, LUNDI), "/?22222222-BbBbB": page(fuseau, MARDI, titre="Atelier")},
              reponses=["O", "O"])
    executer([URL1, "--nom", NOM, "--submit"], m.dep)
    executer(["when2meet.com/?22222222-BbBbB", "--nom", NOM, "--submit"], m.dep)
    contenu = json.loads(open("sondages.json", encoding="utf-8").read())
    assert [e["url"] for e in contenu] == [URL1, URL2]
    assert set(contenu[0]) == {"url", "titre", "nom", "fin", "dernier_envoi"}

    executer(["--liste"], m.dep)
    assert "1. Point projet  [Pierre T]  jusqu'au 02/11/2026" in m.texte

    executer(["--oublier", "1"], m.dep)
    assert [e.url for e in Registre().lire()] == [URL2]
    executer(["--oublier", "https://www.when2meet.com/?22222222-BbBbB#x"], m.dep)
    assert Registre().lire() == []
    assert "Ne sera plus actualisé : Atelier" in m.texte


def test_oublier_sans_argument_demande_le_numero(fuseau):
    Registre().enregistrer(URL1, "Point projet", NOM, 2_000_000_000, 0)
    m = Monde(fuseau, {}, reponses=["1"])
    executer(["--oublier"], m.dep)
    assert Registre().lire() == []


def test_actualiser_plusieurs_sondages(fuseau):
    fin_lundi = fuseau.vers_epoch(datetime(2026, 11, 2, 10))
    Registre().enregistrer(URL1, "Point projet", NOM, fin_lundi, 0)
    Registre().enregistrer(URL2, "Atelier", NOM, fin_lundi + 86400, 0)
    Registre().enregistrer("https://www.when2meet.com/?33333333-CcCcC", "Vieux sondage", NOM, 1_000_000, 0)
    m = Monde(
        fuseau,
        {"/?11111111-AaAaA": page(fuseau, LUNDI, [0, 1, 2, 3]), "/?22222222-BbBbB": page(fuseau, MARDI, [0, 1, 2, 3])},
        rdvs=[(datetime(2026, 11, 2, 9, 30), datetime(2026, 11, 2, 10))],
        reponses=["O"],
    )
    assert executer(["--actualiser", "--submit"], m.dep) in (0, 3)
    assert "Sondage terminé, retiré de la liste : Vieux sondage" in m.texte
    assert "[1/2] Point projet" in m.texte and "[2/2] Atelier" in m.texte
    assert m.envois() == [("11111111", "1100")]
    assert m.questions == ["Envoyer ? (O/N) "]
    assert "Plages libres" not in m.texte  # aperçu condensé en actualisation
    assert [e.url for e in Registre().lire()] == [URL1, URL2]


def test_actualiser_continue_apres_une_erreur(fuseau):
    Registre().enregistrer(URL1, "Cassé", NOM, 2_000_000_000, 0)
    Registre().enregistrer(URL2, "Atelier", NOM, 2_000_000_000, 0)
    m = Monde(fuseau, {"/?11111111-AaAaA": "<html>maintenance</html>",
                       "/?22222222-BbBbB": (page(fuseau, MARDI), page(fuseau, MARDI, [0, 1, 2, 3]))},
              reponses=["O"])
    assert executer(["--actualiser", "--submit"], m.dep) == 1
    assert "ERREUR : Aucun créneau trouvé" in m.texte
    assert m.envois() == [("22222222", "1111")]


def test_actualiser_sans_sondage(fuseau):
    m = Monde(fuseau, {})
    assert executer(["--actualiser", "--submit"], m.dep) == 0
    assert "Aucun sondage à actualiser" in m.texte


def test_noms_echappes_lus_correctement(fuseau):
    epochs = epochs_sondage(fuseau, [LUNDI], time(9), time(10))
    html = page_when2meet(epochs, dispo={"7": [1, 2]}, noms={"7": "Jean d\\'Arc &amp; Co"})
    assert analyser_page(html, URL1).participants == {"Jean d'Arc & Co": frozenset({1, 2})}
