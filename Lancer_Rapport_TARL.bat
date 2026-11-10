@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Rapport TARL

set "PY="
set "PYW="
where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
    set "PYW=pyw -3"
)
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PY=python"
        set "PYW=pythonw"
    )
)
if not defined PY goto nopython

%PY% -c "import openpyxl, pandas, numpy, matplotlib, docx, tkinter, rapidfuzz, xlrd" >nul 2>nul
if errorlevel 1 (
    echo Installation des bibliotheques necessaires - Internet requis cette fois.
    %PY% -m pip install --upgrade pip
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 goto echec
)

start "" %PYW% main.py
exit /b 0

:nopython
echo.
echo  Python n'est pas installe sur cet ordinateur.
echo  Utilisez plutot "Rapport TARL.exe" si vous l'avez, ou installez Python :
echo  https://www.python.org/downloads/  - cochez "Add Python to PATH".
echo.
pause
exit /b 1

:echec
echo.
echo  L'installation a echoue. Verifiez Internet puis relancez.
echo.
pause
exit /b 1
