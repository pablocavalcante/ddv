import sys
import os
import threading
import webbrowser
from streamlit.web import cli as stcli

# Função para encontrar o app.py mesmo quando ele virar um .exe
def resolver_caminho_app():
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, 'app.py')
    return 'app.py'

def abrir_navegador():
    # O Streamlit roda na porta 8501 por padrão
    webbrowser.open_new('http://localhost:8501')

if __name__ == '__main__':
    # 1. Inicia um relógio que espera 2 segundos e abre o navegador
    threading.Timer(2.0, abrir_navegador).start()
    
    # 2. Encontra o seu arquivo principal
    app_path = resolver_caminho_app()
    
    # 3. Engana o sistema simulando que alguém digitou "streamlit run app.py" no terminal
    sys.argv = [
        "streamlit", 
        "run", app_path, 
        "--server.headless", "true", # Impede o Streamlit de tentar abrir abas duplicadas
        "--global.developmentMode", "false"
    ]
    
    # 4. Dá a partida no servidor!
    sys.exit(stcli.main())