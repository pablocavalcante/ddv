import os
import sys
import datetime
import threading
import concurrent.futures
import multiprocessing
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

import customtkinter as ctk
from PIL import Image

# Importações dos serviços
from service_access import gerar_mdb_access
from service_excel import extrair_info_template, processar_arquivo_isolado

# --- FUNÇÃO ESSENCIAL PARA O EXECUTÁVEL (.EXE) ---
def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona para dev e PyInstaller """
    try:
        # O PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Configuração Global do Tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class FrmSelecaoRotina(ctk.CTkFrame):
    def __init__(self, master, callback_sucesso):
        super().__init__(master)
        self.callback_sucesso = callback_sucesso
        self.pack(fill="both", expand=True, padx=20, pady=20)

        self.lbl_titulo = ctk.CTkLabel(self, text="Seleção da Rotina", font=("Roboto", 24, "bold"))
        self.lbl_titulo.pack(pady=(40, 30))

        self.cb_rotinas = ctk.CTkComboBox(self, values=["SJ230133", "SJ071930", "SJ071984"], width=200, height=35)
        self.cb_rotinas.pack(pady=10)
        self.cb_rotinas.set("SJ230133")

        self.btn_ok = ctk.CTkButton(self, text="Confirmar e Iniciar", command=self.confirmar, width=200, height=40)
        self.btn_ok.pack(pady=30)

    def confirmar(self):
        rotina = self.cb_rotinas.get()
        if not rotina: return
        self.pack_forget() 
        self.callback_sucesso(rotina) 

class FrmPrincipal(ctk.CTkFrame):
    def __init__(self, master, rotina_escolhida):
        super().__init__(master)
        self.master = master
        self.rotina_selecionada = rotina_escolhida
        self.pack(fill="both", expand=True)

        # --- CARREGAR ÍCONES USANDO RESOURCE_PATH ---
        pasta_icons = resource_path("icons")

        self.icon_folder = None
        self.icon_play = None
        
        try:
            # Note o uso do resource_path para garantir que o EXE ache a imagem interna
            self.icon_folder = ctk.CTkImage(light_image=Image.open(os.path.join(pasta_icons, "file_open.png")),
                                            dark_image=Image.open(os.path.join(pasta_icons, "file_open.png")),
                                            size=(20, 20))
                                            
            self.icon_play = ctk.CTkImage(light_image=Image.open(os.path.join(pasta_icons, "arrow_circle.png")),
                                          dark_image=Image.open(os.path.join(pasta_icons, "arrow_circle.png")),
                                          size=(24, 24))
        except Exception as e:
            print(f"Aviso: Ícones não encontrados ({e}). O programa rodará sem eles.")

        # --- LAYOUT ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        self.logo = ctk.CTkLabel(self.sidebar, text="DDV\nMigração", font=("Roboto", 20, "bold"))
        self.logo.pack(pady=(30, 30))
        
        self.lbl_info = ctk.CTkLabel(self.sidebar, text=f"Rotina:\n{rotina_escolhida}", font=("Roboto", 14), text_color="gray70")
        self.lbl_info.pack(pady=10)

        self.main_area = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_area.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        self.criar_input_arquivo("Arquivo Header:", "txt_header")
        self.criar_input_arquivo("Arquivo Detail:", "txt_detail")
        self.criar_input_arquivo("Arquivo Índices:", "txt_indices")
        self.criar_input_arquivo("Pasta de Saída:", "txt_output", folder=True)

        self.btn_run = ctk.CTkButton(self.main_area, text="PROCESSAR TUDO", command=self.start, 
                                     height=50, font=("Roboto", 16, "bold"), fg_color="#2CC985", hover_color="#229A65",
                                     image=self.icon_play, compound="right") 
        self.btn_run.pack(fill="x", pady=(30, 20))

        self.pb = ctk.CTkProgressBar(self.main_area)
        self.pb.pack(fill="x", pady=(10, 5))
        self.pb.set(0)
        
        self.lbl_status = ctk.CTkLabel(self.main_area, text="Aguardando início...", font=("Roboto", 12))
        self.lbl_status.pack()

    def criar_input_arquivo(self, label_text, attr_name, folder=False):
        frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=100, anchor="w")
        lbl.pack(side="left")
        
        entry = ctk.CTkEntry(frame, placeholder_text="Selecione o caminho...")
        entry.pack(side="left", fill="x", expand=True, padx=10)
        setattr(self, attr_name, entry) 
        
        btn = ctk.CTkButton(frame, text="", image=self.icon_folder, width=40, 
                            fg_color="#3B8ED0", hover_color="#36719F",
                            command=lambda: self._buscar(entry, folder))
        btn.pack(side="right")

    def _buscar(self, entry_widget, is_folder):
        path = filedialog.askdirectory() if is_folder else filedialog.askopenfilename(filetypes=[("Txt", "*.txt")])
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def start(self): 
        threading.Thread(target=self.run).start()

    def _carregar_indices(self, caminho_indice):
        """Carrega arquivo de índices com tratamento de erros."""
        if not (p_ind_stripped := caminho_indice.strip()) or not os.path.exists(p_ind_stripped):
            return []
        
        indices = []
        with open(p_ind_stripped, 'r', encoding='cp1252') as f:
            for linha in f:
                if len(linha) >= 8:
                    try:
                        dt = datetime.datetime.strptime(linha[:8], "%Y%m%d")
                        val = float(linha[8:].replace(',', '.')) / 100.0
                        indices.append((dt, val))
                    except (ValueError, IndexError):
                        pass
        return indices

    def _carregar_details(self, caminho_details):
        """Carrega arquivo de detalhes e mapeia por processo."""
        with open(caminho_details, 'r', encoding='cp1252') as f:
            linhas = f.readlines()
        
        lines_validas = [l for l in linhas if len(l) >= 120]
        mapa = defaultdict(list)
        for l in lines_validas:
            chave = l[:12].strip() + l[12:21].strip()
            mapa[chave].append(l)
        return lines_validas, mapa

    def run(self):
        try:
            tempo_inicio = datetime.datetime.now()
            self.master.after(0, lambda: self.btn_run.configure(state="disabled", text="PROCESSANDO..."))
            self.master.after(0, lambda: self.lbl_status.configure(text="Iniciando leitura..."))
            
            # Validação de entrada (fail-fast)
            p_head = self.txt_header.get()
            p_det = self.txt_detail.get()
            if not (p_head and p_det):
                raise ValueError("Selecione os arquivos Header e Detail!")
            
            p_ind = self.txt_indices.get()
            p_out = self.txt_output.get()
            
            # Carregar arquivos
            with open(p_head, 'r', encoding='cp1252') as f:
                lines_h = [l for l in f if len(l) >= 180]
            
            idx_list = self._carregar_indices(p_ind)
            dt_lim = idx_list[-1][0] if idx_list else datetime.datetime.now()
            
            lines_d_valid, map_d = self._carregar_details(p_det)

            # --- PROCESSAMENTO DE TEMPLATES ---
            self.master.after(0, lambda: self.lbl_status.configure(text="Gerando Access..."))
            
            p_tpl_xls = resource_path(os.path.join("Templates", "XLS-Matriz.xlsx"))
            p_tpl_mdb = resource_path(os.path.join("Templates", "MDB-Matriz.mdb"))
            
            # Validar templates
            for tpl_path, tpl_name in [(p_tpl_mdb, "MDB"), (p_tpl_xls, "XLS")]:
                if not os.path.exists(tpl_path):
                    raise FileNotFoundError(f"Template {tpl_name} não encontrado: {tpl_path}")

            ok, msg = gerar_mdb_access(lines_h, lines_d_valid, p_out, self.rotina_selecionada, p_tpl_mdb)
            if not ok: 
                self.master.after(0, lambda m=msg: messagebox.showwarning("Aviso Access", m))

            self.master.after(0, lambda: self.lbl_status.configure(text="Processando Excel..."))
            tpl_info = extrair_info_template(p_tpl_xls)
            
            # Construir tarefas de processamento
            tasks = [
                (l, p_tpl_xls, p_out, self.rotina_selecionada, idx_list, map_d.get(l[:12].strip() + l[12:21].strip(), []), dt_lim, tpl_info)
                for l in lines_h
            ]

            tot = len(tasks)
            done = 0
            errs = []
            
            def update_progress(val, total):
                p = val / total
                self.pb.set(p)
                self.lbl_status.configure(text=f"Excel: {val}/{total} ({int(p*100)}%)")

            with concurrent.futures.ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count()-1)) as exe:
                futs = {exe.submit(processar_arquivo_isolado, t): t for t in tasks}
                for f in concurrent.futures.as_completed(futs):
                    res = f.result()
                    done += 1
                    if "ERRO" in res:
                        errs.append(res)
                    self.master.after(0, lambda v=done, t=tot: update_progress(v, t))

            # Gerar relatório final
            tempo_fim = datetime.datetime.now()
            tempo_total = tempo_fim - tempo_inicio
            tempo_str = str(tempo_total).split('.')[0] 

            msg_linhas = [
                f"Processo Finalizado!\n\nTempo Total: {tempo_str}",
                f"Sucesso: {tot - len(errs)}",
                f"Erros: {len(errs)}"
            ]
            if errs:
                msg_linhas.append("\nPrimeiras 5 falhas:")
                msg_linhas.extend(errs[:5])
            
            msg = "\n".join(msg_linhas)
            self.master.after(0, lambda m=msg: messagebox.showinfo("Fim", m))
            self.master.after(0, lambda: self.lbl_status.configure(text=f"Concluído em {tempo_str}"))

        except (FileNotFoundError, ValueError) as e:
            self.master.after(0, lambda erro=str(e): messagebox.showerror("Erro", f"Ocorreu um erro:\n{erro}"))
            self.master.after(0, lambda: self.lbl_status.configure(text="Erro!"))
        except Exception as e:
            self.master.after(0, lambda erro=str(e): messagebox.showerror("Erro (Inesperado)", f"Erro inesperado:\n{erro}"))
            self.master.after(0, lambda: self.lbl_status.configure(text="Erro!"))
        finally:
            self.master.after(0, lambda: self.btn_run.configure(state="normal", text="PROCESSAR TUDO"))