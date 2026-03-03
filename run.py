import os
import sys
import streamlit.web.cli as stcli

if __name__ == "__main__":
    # Descobre a pasta onde o .exe está rodando
    if getattr(sys, 'frozen', False):
        pasta_base = sys._MEIPASS
    else:
        pasta_base = os.path.dirname(os.path.abspath(__file__))
        
    # FORÇA o Windows a trabalhar dentro desta pasta (Corrige caminhos cegos)
    os.chdir(pasta_base)
    
    caminho_app = os.path.join(pasta_base, 'app.py')
    
    # Prepara o comando silencioso para iniciar o sistema
    sys.argv = [
        "streamlit", 
        "run", 
        caminho_app, 
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false"
    ]
    
    # Dá a partida no servidor
    sys.exit(stcli.main())