@echo off
rem =====================================================
rem  Start Extractor (knap-interface) med et dobbeltklik
rem  Laeg denne fil i samme mappe som extractor_gui.py
rem =====================================================

cd /d "%~dp0"

if not exist "extractor_gui.py" (
    echo FEJL: extractor_gui.py ligger ikke i denne mappe.
    echo Laeg .bat-filen sammen med extractor.py, extractor_gui.py og krydstjek.py.
    pause
    exit /b 1
)

rem Brug py-launcheren hvis den findes (folger altid med Python og er versionsuafhaengig)
where py >nul 2>&1
if %errorlevel% == 0 (
    start "" py extractor_gui.py
    exit /b 0
)

rem Reserve: eksplicit sti (ret denne hvis Python er installeret et andet sted)
set "PY=C:\Program Files\Python314\pythonw.exe"
if not exist "%PY%" (
    echo FEJL: Python ikke fundet.
    echo Installer Python fra python.org, eller ret PY-linjen i denne fil.
    pause
    exit /b 1
)
"%PY%" extractor_gui.py
