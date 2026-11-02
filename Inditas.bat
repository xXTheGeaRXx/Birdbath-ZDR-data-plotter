@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" radar_plotter.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo HIBA: A program nem indult el.
    echo Ellenorizze, hogy a .venv mappa letezik-e es a fuggosegek telepitve vannak.
    pause
)
