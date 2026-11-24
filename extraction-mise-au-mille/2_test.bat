@echo off
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py -3
echo Essai sur les 5 derniers jours...
%PY% extraction.py --test
pause
