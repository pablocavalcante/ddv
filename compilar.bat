@echo off
setlocal enabledelayedexpansion

echo ==========================================
echo 1. LIMPANDO AMBIENTE E CACHE
echo ==========================================
call venv\Scripts\activate

if exist "build"     rmdir /s /q "build"
if exist "dist"      rmdir /s /q "dist"
del /f /q *.spec 2>nul
del /f /q "DDV.exe" 2>nul

echo.
echo ==========================================
echo 2. GERANDO EXECUTAVEL (SOLUCAO TKINTER)
echo ==========================================
:: Adicionamos --collect-all tkinter para garantir que o filedialog funcione
:: Adicionamos --hidden-import tkinter.filedialog explicitamente
pyinstaller --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name "DDV" ^
    --add-data "app.py;." ^
    --add-data "access.py;." ^
    --add-data "excel.py;." ^
    --add-data "Templates;Templates/" ^
    --add-data "icons;icons/" ^
    --collect-all streamlit ^
    --collect-all tkinter ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "openpyxl" ^
    --hidden-import "pandas" ^
    --hidden-import "pyodbc" ^
    --exclude-module "matplotlib" ^
    --exclude-module "scipy" ^
    --exclude-module "bokeh" ^
    --exclude-module "seaborn" ^
    --exclude-module "plotly" ^
    --exclude-module "altair" ^
    run.py

echo.
echo ==========================================
echo 3. POS-PROCESSAMENTO E LIMPEZA
echo ==========================================

if exist "dist\DDV.exe" (
    move /y "dist\DDV.exe" "."
    
    :: Limpeza final de pastas para deixar apenas o .exe
    rmdir /s /q "build"
    rmdir /s /q "dist"
    del /f /q "DDV.spec"

    echo.
    echo =============================================
    echo   SUCESSO! Executavel 'DDV.exe' pronto.
    echo =============================================
) else (
    echo.
    echo [ERRO] O executavel nao foi gerado. Verifique os logs acima.
)

pause