@echo off
rem Outlook2Scheduler : remet a jour tous les sondages deja envoyes et encore ouverts.
rem Pour chacun : changements depuis le dernier envoi, puis "Envoyer ? (O/N)".
call "%~dp0o2s.bat" --actualiser
