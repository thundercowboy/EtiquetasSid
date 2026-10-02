@echo off
cd /d "%~dp0"
python etiquetas_sid.py
if errorlevel 1 pause
