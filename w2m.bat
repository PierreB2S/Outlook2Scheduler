@echo off
setlocal
rem Outlook2Scheduler : remplit un sondage When2meet depuis l'agenda Outlook.
rem Double-clic : le lien est lu dans le presse-papiers (sinon il est demande),
rem l'apercu des plages libres s'affiche, puis "Envoyer ? (O/N)" avant tout envoi.
rem En ligne de commande : w2m.bat [lien] [--nom "Prenom N."] [--heures 08:30-18:00] ...

cd /d "%~dp0"
title Outlook2Scheduler

where uv >nul 2>nul
if errorlevel 1 goto :sans_uv

rem Proxy d'entreprise avec inspection SSL : uv utilise les certificats Windows.
rem (UV_SYSTEM_CERTS depuis uv 0.11, UV_NATIVE_TLS pour les versions plus anciennes)
uv run --help 2>nul | findstr /c:"--system-certs" >nul
if errorlevel 1 (set "UV_NATIVE_TLS=1") else (set "UV_SYSTEM_CERTS=1")

rem Si uv ne peut pas telecharger Python (proxy), decommentez et adaptez :
rem set "UV_PYTHON=C:\chemin\vers\python.exe"

uv run w2m %* --submit
set "CODE=%ERRORLEVEL%"
goto :fin

:sans_uv
echo.
echo uv n'est pas installe (ou n'est pas dans le PATH).
echo Pour l'installer, ouvrez PowerShell et lancez :
echo     powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
echo ou, si winget est disponible :
echo     winget install --id=astral-sh.uv -e
echo Puis fermez cette fenetre et relancez w2m.bat.
set "CODE=9"

:fin
echo.
pause
exit /b %CODE%
