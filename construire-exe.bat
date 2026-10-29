@echo off
setlocal
rem Construit la version autonome : dist\Outlook2Scheduler\ et dist\Outlook2Scheduler.zip
rem (o2s.exe + _internal, sans Python ni uv a installer sur le poste cible).

cd /d "%~dp0"
where uv >nul 2>nul
if errorlevel 1 (
    echo uv est necessaire pour construire l'exe : voir README.md.
    set "CODE=9"
    goto :fin
)
uv run --help 2>nul | findstr /c:"--system-certs" >nul
if errorlevel 1 (set "UV_NATIVE_TLS=1") else (set "UV_SYSTEM_CERTS=1")

if exist dist\Outlook2Scheduler rmdir /s /q dist\Outlook2Scheduler
if exist dist\Outlook2Scheduler.zip del /q dist\Outlook2Scheduler.zip

uv run --group build pyinstaller --noconfirm --clean --onedir --console ^
    --exclude-module setuptools --exclude-module pkg_resources ^
    --name o2s --distpath dist --workpath build\pyinstaller --specpath build ^
    packaging\o2s_exe.py
if errorlevel 1 (
    set "CODE=1"
    goto :fin
)

move dist\o2s dist\Outlook2Scheduler >nul
copy /y config.example.toml dist\Outlook2Scheduler\ >nul
copy /y packaging\o2s-actualiser.bat dist\Outlook2Scheduler\ >nul
copy /y packaging\LISEZMOI.txt dist\Outlook2Scheduler\ >nul
copy /y LICENSE dist\Outlook2Scheduler\LICENSE.txt >nul

powershell -NoProfile -Command "Compress-Archive -Path 'dist\Outlook2Scheduler' -DestinationPath 'dist\Outlook2Scheduler.zip' -Force"
echo.
echo Pret : dist\Outlook2Scheduler.zip
set "CODE=0"

:fin
echo.
pause
exit /b %CODE%
