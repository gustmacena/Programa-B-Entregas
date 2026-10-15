@echo off
rem Gera UM UNICO arquivo: dist\ConsultaLegado.exe
rem Rode este .bat SOMENTE na maquina que tem Python (a sua). O usuario final so recebe o .exe.
cd /d "%~dp0"
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm --onefile --windowed --name ConsultaLegado --icon icone.ico --add-data "icone.ico;." --add-data "logo;logo" --collect-all PySide6.QtPdf --hidden-import python_calamine app.py
echo.
echo Pronto! Entregue ao usuario apenas o arquivo: dist\ConsultaLegado.exe
echo (nao precisa de Python nem de nenhum .bat na maquina dele)
pause
