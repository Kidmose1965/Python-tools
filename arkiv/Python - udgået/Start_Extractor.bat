@echo off
rem =====================================================
rem  Start Extractor (knap-interface) med ét dobbeltklik
rem  Laeg denne fil i samme mappe som extractor_gui.py
rem =====================================================

set "PY=C:\Users\BKI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

rem Skift til mappen hvor denne .bat-fil ligger
cd /d "%~dp0"

if not exist "%PY%" (
    echo FEJL: Python blev ikke fundet paa den forventede sti:
    echo   %PY%
    echo.
    echo Stien kan vaere aendret efter en opdatering. Find den nye med:
    echo   dir C:\python.exe /s /b
    echo og ret PY-linjen oeverst i denne fil.
    pause
    exit /b 1
)

if not exist "extractor_gui.py" (
    echo FEJL: extractor_gui.py ligger ikke i denne mappe.
    echo Laeg .bat-filen sammen med extractor.py og extractor_gui.py.
    pause
    exit /b 1
)

"%PY%" extractor_gui.py

if errorlevel 1 (
    echo.
    echo Programmet sluttede med en fejl - se beskeden ovenfor.
    pause
)
