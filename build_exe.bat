@echo off
rem Gera a pasta dist\ConsultaLegado com o ConsultaLegado.exe (nao precisa de Python para usar)
cd /d "%~dp0"
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm --windowed --name ConsultaLegado --icon icone.ico --add-data "icone.ico;." --add-data "logo;logo" --collect-all PySide6.QtPdf --hidden-import python_calamine app.py
echo.
echo Pronto! Copie a pasta dist\ConsultaLegado para onde quiser e abra o ConsultaLegado.exe
pause
