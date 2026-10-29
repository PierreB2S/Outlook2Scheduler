"""Point d'entrée de o2s.exe (version autonome, sans uv ni Python installé).

Se comporte comme o2s.bat : travaille dans le dossier de l'exe (config.toml et
sondages.json à côté), envoie après aperçu et « Envoyer ? (O/N) », et garde la
fenêtre ouverte à la fin quand on l'a lancé par double-clic.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from outlook2scheduler.cli import main


def dossier_programme() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def arguments(argv: list[str]) -> list[str]:
    """Ajoute --submit comme o2s.bat (sans --submit, rien ne serait envoyé)."""
    return argv if "--submit" in argv else [*argv, "--submit"]


def lance_par_double_clic() -> bool:
    """Vrai si la console n'appartient qu'à ce programme (double-clic), faux depuis cmd ou un .bat."""
    try:
        import ctypes

        processus = (ctypes.c_uint * 4)()
        return ctypes.windll.kernel32.GetConsoleProcessList(processus, 4) == 1
    except (AttributeError, OSError):
        return False


def principal() -> int:
    os.chdir(dossier_programme())
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleTitleW("Outlook2Scheduler")
    except (AttributeError, OSError):
        pass
    code = main(arguments(sys.argv[1:]))
    if lance_par_double_clic():
        try:
            input("\nAppuyez sur Entrée pour fermer...")
        except (EOFError, KeyboardInterrupt):
            pass
    return code


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(principal())
