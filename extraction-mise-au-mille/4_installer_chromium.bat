@echo off
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py -3
echo Telechargement du navigateur Chromium reserve au bot, environ 150 Mo...
%PY% -m playwright install chromium
if errorlevel 1 (
  echo.
  echo ECHEC : le telechargement est sans doute bloque par le reseau de l'entreprise.
) else (
  echo.
  echo Termine. Relance maintenant 2_test.bat
)
pause
