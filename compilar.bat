@echo off

echo ==========================================
echo ATIVANDO AMBIENTE VIRTUAL...
echo ==========================================
call venv\Scripts\activate

echo.
echo ==========================================
echo LIMPANDO CACHE E CORRIGINDO PYARROW...
echo ==========================================
if exist "%appdata%\pyinstaller" rmdir /s /q "%appdata%\pyinstaller"
if exist "venv\Lib\site-packages\pyarrow\include" rmdir /s /q "venv\Lib\site-packages\pyarrow\include"
if exist "build"    rmdir /s /q "build"
if exist "dist"     rmdir /s /q "dist"
if exist "DDV_App"  rmdir /s /q "DDV_App"
del /f /q *.spec 2>nul

echo.
echo ==========================================
echo INICIANDO EMPACOTAMENTO...
echo ==========================================
pyinstaller --noconfirm --clean ^
    --onedir ^
    --windowed ^
    --name "DDV" ^
    --add-data "app.py;." ^
    --add-data "access.py;." ^
    --add-data "excel.py;." ^
    --add-data "Templates;Templates/" ^
    --add-data "icons;icons/" ^
    --collect-all streamlit ^
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
echo FINALIZANDO E LIMPANDO...
echo ==========================================
if exist "dist\DDV\DDV.exe" (
    xcopy /e /i /y "dist\DDV" "DDV_App\"
    rmdir /s /q "build"
    rmdir /s /q "dist"
    del /f /q "DDV.spec"
    echo.
    echo =============================================
    echo  SUCESSO! Distribua a pasta DDV_App\ inteira
    echo  Execute: DDV_App\DDV.exe
    echo =============================================
) else (
    echo [ERRO] O executavel nao foi gerado. Verifique o log acima.
)

pause