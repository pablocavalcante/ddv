import os
import sys
import multiprocessing
import customtkinter as ctk
from interface import FrmSelecaoRotina, FrmPrincipal, resource_path

# --- FUNÇÃO AUXILIAR PARA CENTRALIZAR ---
def centralizar_janela(janela, largura, altura):
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    pos_x = (largura_tela - largura) // 2
    pos_y = (altura_tela - altura) // 2
    janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

# Note que agora recebemos 'janela_raiz' como argumento
def iniciar_app_principal(janela_raiz, rotina):
    centralizar_janela(janela_raiz, 600, 450)    
    app_principal = FrmPrincipal(janela_raiz, rotina)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    
    root = ctk.CTk()
    root.title("DDV")
    
    # Configurar o ícone da janela
    try:
        icon_path = resource_path(os.path.join("icons", "msaccess.ico"))
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"Aviso: Ícone não encontrado ({e})")
    
    centralizar_janela(root, 400, 300)

    
    def callback_bridge(rotina):
        iniciar_app_principal(root, rotina)

    selecao = FrmSelecaoRotina(root, callback_bridge)
    
    root.mainloop()