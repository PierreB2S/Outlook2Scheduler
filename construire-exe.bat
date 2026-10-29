@echo off
setlocal
rem Construit la version autonome : dist\Outlook2Scheduler\ et dist\Outlook2Scheduler.zip
rem (o2s.exe + _internal, sans Python ni uv a installer sur le poste cible).
rem
rem PyInstaller et les dependances sont installes dans un environnement
rem temporaire, a partir de l'index de paquets configure pour uv sur ce poste
rem (PyPI, ou depot interne). uv.lock et .venv ne sont pas touches.

cd /d "%~dp0"
title Outlook2Scheduler - construction de l'exe
where uv >nul 2>nul
if errorlevel 1 (
    echo uv est necessaire pour construire l'exe : voir README.md.
    set "CODE=9"
    goto :fin
)
uv run --help 2>nul | findstr /c:"--system-certs" >nul
if errorlevel 1 (set "UV_NATIVE_TLS=1") else (set "UV_SYSTEM_CERTS=1")
rem Depot de paquets lent : on laisse plus de temps a chaque requete.
if not defined UV_HTTP_TIMEOUT set "UV_HTTP_TIMEOUT=120"

rem Si uv ne trouve pas Python 3.11 ou plus, decommentez et adaptez :
rem set "UV_PYTHON=C:\chemin\vers\python.exe"

if exist dist\Outlook2Scheduler rmdir /s /q dist\Outlook2Scheduler
if exist dist\Outlook2Scheduler.zip del /q dist\Outlook2Scheduler.zip

uv run --no-project --python ">=3.11" ^
    --with "pyinstaller>=6.10" --with "pywin32>=306" --with "requests>=2.31" --with "truststore>=0.9" ^
    pyinstaller --noconfirm --clean --onedir --console ^
    --paths src --exclude-module setuptools --exclude-module pkg_resources ^
    --name o2s --distpath dist --workpath build\pyinstaller --specpath build ^
    packaging\o2s_exe.py
if errorlevel 1 goto :echec

move dist\o2s dist\Outlook2Scheduler >nul
copy /y config.example.toml dist\Outlook2Scheduler\ >nul
copy /y packaging\o2s-actualiser.bat dist\Outlook2Scheduler\ >nul
copy /y packaging\LISEZMOI.txt dist\Outlook2Scheduler\ >nul
copy /y LICENSE dist\Outlook2Scheduler\LICENSE.txt >nul

powershell -NoProfile -Command "Compress-Archive -Path 'dist\Outlook2Scheduler' -DestinationPath 'dist\Outlook2Scheduler.zip' -Force"
echo.
echo Pret : dist\Outlook2Scheduler.zip
set "CODE=0"
goto :fin

:echec
echo.
echo La construction a echoue.
echo Si l'erreur parle de telechargement ("Failed to fetch", "tcp connect error") :
echo uv n'a pas pu joindre le depot de paquets. Elements utiles :
echo.
echo  - Variables proxy et uv de cette fenetre :
set | findstr /i /b "HTTP_PROXY HTTPS_PROXY ALL_PROXY NO_PROXY UV_" 2>nul
echo.
echo  - Si le depot est interne et qu'un proxy est defini ci-dessus, il faut
echo    peut-etre ajouter son nom a NO_PROXY (sinon uv passe par le proxy).
echo  - Test simple, dans une invite de commandes :  uvx pyinstaller --version
set "CODE=1"

:fin
echo.
pause
exit /b %CODE%
