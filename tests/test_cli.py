"""Scénarios de bout en bout : faux Outlook + faux When2meet, vraie logique."""

from datetime import date, datetime, time

import pytest
from conftest import FausseSession, FauxFreeBusy, epochs_sondage, page_when2meet

from outlook2scheduler.cibles.when2meet import When2meet
from outlook2scheduler.cli import Dependances, executer, main
from outlook2scheduler.erreurs import ErreurConfig, ErreurSondageJoursSemaine
from outlook2scheduler.sources.base import OCCUPE, PROVISOIRE
from outlook2scheduler.sources.outlook import SourceOutlook

URL = "https://www.when2meet.com/?12345678-AbCdE"


class Monde:
    """Assemble les doublures et capture la sortie."""

    def __init__(self, fuseau, epochs, evenements=(), reponses=(), semantique="ecoule", presse_papiers=""):
        self.session = FausseSession()
        self.session.programmer("/?12345678-AbCdE", page_when2meet(epochs))
        self.session.programmer("/ProcessLogin.php", "4242")
        self.session.programmer("/SaveTimes.php", "")
        self.client = FauxFreeBusy(fuseau, list(evenements), semantique=semantique)
        self.sorties: list[str] = []
        self.questions: list[str] = []
        self._reponses = list(reponses)
        self.dep = Dependances(
            fabrique_source=lambda f: SourceOutlook(self.client, f),
            fabrique_cibles=lambda: [When2meet(self.session)],
            fuseau=fuseau,
            entree=self._entree,
            sortie=self.sorties.append,
            presse_papiers=lambda: presse_papiers,
        )

    def _entree(self, question):
        self.questions.append(question)
        return self._reponses.pop(0)

    @property
    def texte(self):
        return "\n".join(self.sorties)

    def envois(self):
        return self.session.posts("/SaveTimes.php")


def test_apercu_sans_envoi(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(12))
    m = Monde(fuseau, epochs, [(datetime(2026, 11, 2, 10), datetime(2026, 11, 2, 10, 30), OCCUPE)])
    assert executer([URL], m.dep) == 0
    assert "lun. 02/11/2026 : 09:00-10:00, 10:30-12:00" in m.texte
    assert "rien n'a été envoyé" in m.texte
    assert m.envois() == []
    assert m.session.posts("/ProcessLogin.php") == []


@pytest.mark.parametrize("reponse,envoye", [("N", False), ("o", True), ("", False)])
def test_confirmation_avant_envoi(fuseau, tmp_path, monkeypatch, reponse, envoye):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs, reponses=[reponse])
    executer([URL, "--nom", "Testeur", "--submit"], m.dep)
    assert m.questions == ["Envoyer ? (O/N) "]
    assert bool(m.envois()) is envoye


def test_nom_demande_si_absent(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs, reponses=["Testeur", "O"])
    executer([URL, "--submit"], m.dep)
    assert m.session.posts("/ProcessLogin.php")[0]["name"] == "Testeur"


def test_lien_lu_dans_le_presse_papiers(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs, presse_papiers=f"  {URL}\r\n")
    assert executer([], m.dep) == 0
    assert m.questions == []
    assert "presse-papiers" in m.texte


def test_lien_demande_si_presse_papiers_sans_lien(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs, reponses=["pas un lien", f'"{URL}"'], presse_papiers="du texte quelconque")
    assert executer([], m.dep) == 0
    assert len(m.questions) == 2


@pytest.mark.parametrize("semantique", ["ecoule", "murale"])
def test_sondage_a_cheval_sur_le_changement_d_heure(fuseau, tmp_path, monkeypatch, semantique):
    monkeypatch.chdir(tmp_path)
    jours = [date(2026, 10, 23), date(2026, 10, 24), date(2026, 10, 25), date(2026, 10, 26)]
    epochs = epochs_sondage(fuseau, jours, time(9), time(11))
    evenements = [
        (datetime(2026, 10, 23, 9, 30), datetime(2026, 10, 23, 10), OCCUPE),
        (datetime(2026, 10, 26, 10), datetime(2026, 10, 26, 10, 30), OCCUPE),
    ]
    m = Monde(fuseau, epochs, evenements, reponses=["O"], semantique=semantique)
    executer([URL, "--nom", "Testeur", "--submit", "--jours-exclus", "sam"], m.dep)
    dispo = m.envois()[0]["availability"]
    vendredi, samedi, dimanche, lundi = (dispo[i * 8:(i + 1) * 8] for i in range(4))
    assert vendredi == "11001111"
    assert samedi == "00000000"
    assert lundi == "11110011"
    assert "dim. 25/10/2026 est un jour de changement d'heure" in m.texte
    if semantique == "ecoule":
        assert dimanche == "11111111"


def test_dates_non_consecutives(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    jours = [date(2026, 11, 2), date(2026, 11, 12), date(2026, 12, 1)]
    epochs = epochs_sondage(fuseau, jours, time(9), time(10))
    m = Monde(fuseau, epochs, [(datetime(2026, 11, 12, 9), datetime(2026, 11, 12, 9, 15), PROVISOIRE)], reponses=["O"])
    executer([URL, "--nom", "T", "--submit"], m.dep)
    assert m.envois()[0]["availability"] == "1111" + "0111" + "1111"
    assert len(m.client.appels) == 3
    m2 = Monde(fuseau, epochs, [(datetime(2026, 11, 12, 9), datetime(2026, 11, 12, 9, 15), PROVISOIRE)], reponses=["O"])
    executer([URL, "--nom", "T", "--submit", "--provisoire-libre"], m2.dep)
    assert m2.envois()[0]["availability"] == "1" * 12


def test_sondage_jours_de_la_semaine_refuse(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    m = Monde(fuseau, [279_795_600 + 900 * i for i in range(4)])
    with pytest.raises(ErreurSondageJoursSemaine):
        executer([URL], m.dep)
    assert m.client.appels == []


def test_configuration(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "config.toml").write_text(
        'nom = "Depuis Config"\nheures = "09:30-10:00"\nmarge_minutes = 0\n', encoding="utf-8"
    )
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs, reponses=["O"])
    executer([URL, "--submit"], m.dep)
    assert m.session.posts("/ProcessLogin.php")[0]["name"] == "Depuis Config"
    assert m.envois()[0]["availability"] == "0011"


def test_configuration_cle_inconnue(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "config.toml").write_text('nmo = "faute de frappe"\n', encoding="utf-8")
    with pytest.raises(ErreurConfig, match="nmo"):
        executer([URL], Monde(fuseau, []).dep)


def test_diagnostic_agenda(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    m = Monde(fuseau, [], [(datetime(2026, 11, 2, 10, 15), datetime(2026, 11, 2, 10, 45), OCCUPE)])
    assert executer(["--agenda", "2026-11-02"], m.dep) == 0
    assert "  10:00  .XX." in m.texte
    assert m.session.requetes == []


def test_main_affiche_les_erreurs_sans_trace(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    assert main(["https://exemple.invalid/pas-un-sondage"]) == 1
    err = capsys.readouterr().err
    assert "ERREUR" in err and "Traceback" not in err


def test_config_exemple_valide(tmp_path):
    from pathlib import Path

    from outlook2scheduler.config import charger

    exemple = Path(__file__).parent.parent / "config.example.toml"
    config = charger(exemple)
    assert config.heures == "08:30-18:00"
    assert config.jours_exclus == ["sam", "dim"]
