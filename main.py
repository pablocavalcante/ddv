import multiprocessing
import customtkinter as ctk
from interface import FrmSelecaoRotina, FrmPrincipal

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
    
    centralizar_janela(root, 400, 300)
    
    # Passamos 'root' junto com a função usando lambda para ficar elegante
    # Ou altere no interface.py para passar apenas a rotina e a função pegar o root global
    # Mas a forma mais simples sem mexer no interface.py é:
    
    def callback_bridge(rotina):
        iniciar_app_principal(root, rotina)

    selecao = FrmSelecaoRotina(root, callback_bridge)
    
    root.mainloop()