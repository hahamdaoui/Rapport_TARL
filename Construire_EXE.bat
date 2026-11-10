@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Construction de Rapport TARL.exe

echo ==================================================
echo   Construction de "Rapport TARL.exe"
echo   (a faire une seule fois, sur UN PC avec Python)
echo ==================================================
echo.

set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY goto nopython

echo [1/3] Installation des outils - Internet requis...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto echec

echo.
echo [2/3] Creation de l'executable - 2 a 5 minutes, patience...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Rapport TARL" ^
  --icon "assets\rapport_tarl.ico" ^
  --collect-data docx ^
  --paths . ^
  --hidden-import matplotlib.backends.backend_agg ^
  --hidden-import rapidfuzz ^
  --hidden-import xlrd ^
  --exclude-module IPython --exclude-module pytest --exclude-module scipy ^
  --exclude-module PyQt5 --exclude-module PyQt6 --exclude-module PySide2 --exclude-module PySide6 ^
  --exclude-module notebook --exclude-module sphinx ^
  main.py
if errorlevel 1 goto echec

echo.
echo [3/3] Preparation du dossier a distribuer...
if not exist Distribution mkdir Distribution
copy /y "dist\Rapport TARL.exe" "Distribution\Rapport TARL.exe" >nul
if exist synonymes_utilisateur.json copy /y synonymes_utilisateur.json "Distribution\synonymes_utilisateur.json" >nul
if exist build rmdir /s /q build
if exist "Rapport TARL.spec" del /q "Rapport TARL.spec"

echo.
echo ==================================================
echo   TERMINE !
echo   Le fichier a distribuer se trouve dans :
echo   %~dp0Distribution\Rapport TARL.exe
echo ==================================================
explorer "%~dp0Distribution"
pause
exit /b 0

:nopython
echo.
echo  Python est introuvable sur ce PC.
echo  Installez-le depuis https://www.python.org/downloads/
echo  en cochant "Add Python to PATH", puis relancez ce fichier.
echo.
pause
exit /b 1

:echec
echo.
echo  Une etape a echoue. Lisez le message ci-dessus, verifiez Internet
echo  et relancez ce fichier.
echo.
pause
exit /b 1
