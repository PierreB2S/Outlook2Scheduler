"""Point d'entrée de o2s.exe : même comportement que o2s.bat, dans le dossier de l'exe."""

import sys

from outlook2scheduler import exe


def test_ajoute_submit_comme_le_bat():
    assert exe.arguments([]) == ["--submit"]
    assert exe.arguments(["--actualiser"]) == ["--actualiser", "--submit"]
    assert exe.arguments(["lien", "--submit"]) == ["lien", "--submit"]


def test_dossier_de_l_exe_quand_compile(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "o2s.exe"))
    assert exe.dossier_programme() == tmp_path.resolve()


def test_travaille_dans_le_dossier_de_l_exe(monkeypatch, tmp_path):
    recu = {}
    monkeypatch.setattr(exe, "dossier_programme", lambda: tmp_path)
    monkeypatch.setattr(exe, "lance_par_double_clic", lambda: False)
    monkeypatch.setattr(sys, "argv", ["o2s.exe", "--liste"])

    def faux_main(argv):
        import os

        recu["argv"], recu["dossier"] = argv, os.getcwd()
        return 0

    monkeypatch.setattr(exe, "main", faux_main)
    monkeypatch.chdir(tmp_path.parent)
    assert exe.principal() == 0
    assert recu == {"argv": ["--liste", "--submit"], "dossier": str(tmp_path)}
