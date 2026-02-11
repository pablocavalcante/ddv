@echo off
echo Iniciando a criacao do executavel DDV...
echo Aguarde, isso pode levar alguns minutos...

pyinstaller --noconfirm --onefile --windowed --name "DDV" --add-data "icons;icons" --add-data "Templates;Templates" --icon "icons/msaccess.ico" main.py

echo.
echo ===========================================
echo Processo concluido! O arquivo DDV.exe esta na pasta 'dist'.
echo ===========================================
pause