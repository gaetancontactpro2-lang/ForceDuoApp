@echo off
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py -3
echo Assemblage des fichiers Mise au mille en un seul fichier Excel...
%PY% assembler.py %*
pause
