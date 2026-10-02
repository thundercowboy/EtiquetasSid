@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --onefile --windowed --name EtiquetasSid --icon "imgs\icon.png" --add-data "imgs\icon.png;imgs" etiquetas_sid.py
if errorlevel 1 exit /b 1
echo Copie dist\EtiquetasSid.exe para esta pasta, ao lado de BD e responsaveis.txt.
pause
