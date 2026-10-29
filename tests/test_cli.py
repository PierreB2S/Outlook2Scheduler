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
            fabrique_source=lambda f, source: SourceOutlook(self.client, f),
            fabrique_cibles=lambda mode: [When2meet(self.session)],
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


@pytest.mark.parametrize(
    "heures", ['"11:00-12:00, 14:00-15:00"', '["14:00-15:00", "11:00-12:00"]']
)
def test_configuration_pause_de_midi(fuseau, tmp_path, monkeypatch, heures):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "config.toml").write_text(f'nom = "N"\nheures = {heures}\n', encoding="utf-8")
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(11), time(15))  # 16 créneaux
    m = Monde(fuseau, epochs, reponses=["O"])
    executer([URL, "--submit"], m.dep)
    assert m.envois()[0]["availability"] == "1111" + "0" * 8 + "1111"


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


@pytest.mark.parametrize(
    "contenu,argv,attendu",
    [("", [], "windows"), ('reseau = "python"\n', [], "python"), ('reseau = "python"\n', ["--reseau", "windows"], "windows")],
)
def test_choix_du_mode_reseau(fuseau, tmp_path, monkeypatch, contenu, argv, attendu):
    import sys

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "platform", "win32")
    (tmp_path / "config.toml").write_text(contenu, encoding="utf-8")
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m = Monde(fuseau, epochs)
    modes = []
    m.dep.fabrique_cibles = lambda mode: modes.append(mode) or [When2meet(m.session)]
    executer([URL, *argv], m.dep)
    assert modes == [attendu]


def _monde_rdv(fuseau, epochs, rdvs, reponses=()):
    from outlook2scheduler.sources.outlook_rdv import RendezVous, SourceOutlookRendezVous

    class Client:
        def rendez_vous(self, debut, fin):
            return [RendezVous(fuseau.vers_epoch(d), fuseau.vers_epoch(f), s) for d, f, s in rdvs]

    m = Monde(fuseau, epochs, reponses=reponses)
    sources = []
    m.dep.fabrique_source = lambda f, source: sources.append(source) or SourceOutlookRendezVous(Client(), f)
    return m, sources


def test_source_rendez_vous_par_defaut_de_bout_en_bout(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m, sources = _monde_rdv(fuseau, epochs, [(datetime(2026, 11, 2, 9, 15), datetime(2026, 11, 2, 9, 45), OCCUPE)], ["O"])
    executer([URL, "--nom", "T", "--submit"], m.dep)
    assert sources == ["rendez-vous"]
    assert m.envois()[0]["availability"] == "1001"
    assert "Agenda  : 1 rendez-vous" in m.texte


def test_alerte_si_aucun_rendez_vous_lu(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m, _ = _monde_rdv(fuseau, epochs, [])
    executer([URL], m.dep)
    assert "Aucun rendez-vous lu dans Outlook" in m.texte


def test_source_freebusy_au_choix(fuseau, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(10))
    m, sources = _monde_rdv(fuseau, epochs, [])
    executer([URL, "--source", "freebusy"], m.dep)
    assert sources == ["freebusy"]


def test_config_avec_bom_du_bloc_notes(tmp_path):
    from outlook2scheduler.config import charger

    fichier = tmp_path / "config.toml"
    fichier.write_text('heures = "08:00-12:00, 14:00-18:00"\n', encoding="utf-8-sig")
    assert charger(fichier).heures == "08:00-12:00, 14:00-18:00"


@pytest.mark.parametrize("argv,alerte", [(["--source", "freebusy"], True), ([], False)])
def test_alerte_teletravail_avec_freebusy(fuseau, tmp_path, monkeypatch, argv, alerte):
    # Bloc télétravail 9h-12h contenant une réunion 10h-10h30. Le faux Outlook
    # fusionne comme le redoute la macro VBA : « travaille ailleurs » masque la réunion.
    from outlook2scheduler.sources.base import AILLEURS

    monkeypatch.chdir(tmp_path)
    epochs = epochs_sondage(fuseau, [date(2026, 11, 2)], time(9), time(12))
    m = Monde(fuseau, epochs, [
        (datetime(2026, 11, 2, 9), datetime(2026, 11, 2, 12), AILLEURS),
        (datetime(2026, 11, 2, 10), datetime(2026, 11, 2, 10, 30), OCCUPE),
    ])
    executer([URL, *argv], m.dep)
    assert ("12 créneau(x) libre(s) dans des blocs « travaille ailleurs »" in m.texte) is alerte
