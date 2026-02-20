@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===================================================
echo             SISTEMA DDV - INICIALIZACAO
echo ===================================================

:: Verifica se o ambiente virtual (venv) já existe na máquina do cliente
IF NOT EXIST "venv\Scripts\activate.bat" (
    echo [Primeiro Uso] Preparando o sistema para rodar...
    echo Isso pode levar alguns minutos.
    
    :: Cria o venv na máquina do cliente
    python -m venv venv
    
    :: Ativa e instala as bibliotecas
    call venv\Scripts\activate.bat
    echo Instalando dependencias necessarias...
    pip install -r requirements.txt
    
    echo Pronto! Sistema configurado.
) ELSE (
    :: Se já existir, apenas ativa (segunda vez em diante será bem rápido)
    call venv\Scripts\activate.bat
)

echo Iniciando o aplicativo no navegador...
echo ===================================================

:: --- NOVIDADE AQUI ---
:: Configura o Streamlit para NÃO pedir e-mail e NÃO coletar dados do usuário
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_GLOBAL_SHOW_EMAIL_PROMPT=false

streamlit run app.py --logger.level=warning

pause