@echo off
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py -3
echo Installation de Playwright, a faire une seule fois...
%PY% -m pip install --upgrade playwright
if errorlevel 1 (
  echo.
  echo ECHEC : verifie que Python est installe et que pip a acces a Internet.
) else (
  echo.
  echo Installation terminee. Lance maintenant 2_test.bat
)
pause
