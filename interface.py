import os
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

# Configuração Global do Tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class FrmSelecaoRotina(ctk.CTkFrame):
    def __init__(self, master, callback_sucesso):
        super().__init__(master)
        self.callback_sucesso = callback_sucesso
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        self.lbl_titulo = ctk.CTkLabel(self, text="Seleção da Rotina", font=("Roboto", 24, "bold"))
        self.lbl_titulo.pack(pady=(40, 30))

        self.cb_rotinas = ctk.CTkComboBox(self, values=["SJ230133", "SJ071930", "SJ071984"], width=200, height=35)
        self.cb_rotinas.pack(pady=10)
        self.cb_rotinas.set("SJ230133")

        # Botão
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

        # --- CARREGAR ÍCONES ---
        pasta_app = os.path.dirname(__file__)
        pasta_icons = os.path.join(pasta_app, "icons")

        self.icon_folder = None
        self.icon_play = None
        
        try:
            self.icon_folder = ctk.CTkImage(light_image=Image.open(os.path.join(pasta_icons, "file_open.png")),
                                            dark_image=Image.open(os.path.join(pasta_icons, "file_open.png")),
                                            size=(20, 20))
                                            
            self.icon_play = ctk.CTkImage(light_image=Image.open(os.path.join(pasta_icons, "arrow_circle.png")),
                                          dark_image=Image.open(os.path.join(pasta_icons, "arrow_circle.png")),
                                          size=(24, 24))
        except Exception:
            print("Aviso: Ícones não encontrados. O programa rodará sem eles.")

        # --- LAYOUT ---
        
        # Sidebar (Lateral Esquerda)
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        self.logo = ctk.CTkLabel(self.sidebar, text="DDV\nMigração", font=("Roboto", 20, "bold"))
        self.logo.pack(pady=(30, 30))
        
        self.lbl_info = ctk.CTkLabel(self.sidebar, text=f"Rotina:\n{rotina_escolhida}", font=("Roboto", 14), text_color="gray70")
        self.lbl_info.pack(pady=10)

        # Área Principal (Direita)
        self.main_area = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_area.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        # Inputs de Arquivo
        self.criar_input_arquivo("Arquivo Header:", "txt_header")
        self.criar_input_arquivo("Arquivo Detail:", "txt_detail")
        self.criar_input_arquivo("Arquivo Índices:", "txt_indices")
        self.criar_input_arquivo("Pasta de Saída:", "txt_output", folder=True)

        # Botão Processar
        self.btn_run = ctk.CTkButton(self.main_area, text="PROCESSAR TUDO", command=self.start, 
                                     height=50, font=("Roboto", 16, "bold"), fg_color="#2CC985", hover_color="#229A65",
                                     image=self.icon_play, compound="right") 
        self.btn_run.pack(fill="x", pady=(30, 20))

        # Barra de Progresso e Status
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

    def start(self): threading.Thread(target=self.run).start()

    def run(self):
        try:
            print(">>> INICIANDO O PROCESSO...") 
            
            # --- 1. MARCA A HORA DE INÍCIO ---
            tempo_inicio = datetime.datetime.now()
            
            # Travar interface
            self.master.after(0, lambda: self.btn_run.configure(state="disabled", text="PROCESSANDO..."))
            self.master.after(0, lambda: self.lbl_status.configure(text="Iniciando leitura..."))
            
            # Paths
            p_head = self.txt_header.get()
            p_det = self.txt_detail.get()
            p_ind = self.txt_indices.get()
            p_out = self.txt_output.get()
            
            if not p_head or not p_det:
                raise Exception("Selecione os arquivos Header e Detail!")

            print(f"Lendo Header: {p_head}")
            with open(p_head, 'r', encoding='cp1252') as f: lines_h = [l for l in f if len(l)>=180]
            
            # Indices
            idx_list = []
            if os.path.exists(p_ind) and p_ind.strip() != "":
                with open(p_ind, 'r', encoding='cp1252') as f:
                    for l in f:
                        if len(l)>=8:
                            try: idx_list.append((datetime.datetime.strptime(l[:8], "%Y%m%d"), float(l[8:].replace(',', '.'))/100.0))
                            except: pass
            dt_lim = idx_list[-1][0] if idx_list else datetime.datetime.now()

            # Detail
            print(f"Lendo Detail: {p_det}")
            with open(p_det, 'r', encoding='cp1252') as f: lines_d_raw = f.readlines()
            
            map_d = defaultdict(list)
            lines_d_valid = []
            for l in lines_d_raw:
                if len(l)>=120:
                    lines_d_valid.append(l)
                    map_d[l[:12].strip()+l[12:21].strip()].append(l)
            del lines_d_raw

            # ACCESS
            self.master.after(0, lambda: self.lbl_status.configure(text="Gerando Access..."))
            base_dir = os.getcwd()
            p_tpl_xls = os.path.join(base_dir, "Templates", "XLS-Matriz.xlsx")
            p_tpl_mdb = os.path.join(base_dir, "Templates", "MDB-Matriz.mdb")
            
            if not os.path.exists(p_tpl_mdb): raise Exception(f"Template MDB não achado: {p_tpl_mdb}")

            ok, msg = gerar_mdb_access(lines_h, lines_d_valid, p_out, self.rotina_selecionada, p_tpl_mdb)
            if not ok: 
                print(f"Erro Access: {msg}")
                self.master.after(0, lambda m=msg: messagebox.showwarning("Aviso Access", m))

            # EXCEL
            self.master.after(0, lambda: self.lbl_status.configure(text="Processando Excel..."))
            
            if not os.path.exists(p_tpl_xls): raise Exception(f"Template XLS não achado: {p_tpl_xls}")
            tpl_info = extrair_info_template(p_tpl_xls)
            
            tasks = []
            for l in lines_h:
                k = l[:12].strip() + l[12:21].strip()
                tasks.append((l, p_tpl_xls, p_out, self.rotina_selecionada, idx_list, map_d.get(k, []), dt_lim, tpl_info))

            tot = len(tasks)
            print(f"Iniciando Multiprocessamento de {tot} arquivos...")
            
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
                    if "ERRO" in res: errs.append(res)
                    self.master.after(0, lambda v=done, t=tot: update_progress(v, t))

            # --- 2. CALCULA O TEMPO FINAL ---
            tempo_fim = datetime.datetime.now()
            tempo_total = tempo_fim - tempo_inicio
            
            # Formata para ficar bonito (tira os milissegundos extras)
            tempo_str = str(tempo_total).split('.')[0] 

            msg = f"Processo Finalizado!\n\nTempo Total: {tempo_str}\nSucesso: {tot-len(errs)}\nErros: {len(errs)}"
            if errs: msg += "\n" + "\n".join(errs[:5])
            
            print(f"Concluído em {tempo_str}")
            self.master.after(0, lambda m=msg: messagebox.showinfo("Fim", m))
            self.master.after(0, lambda: self.lbl_status.configure(text=f"Concluído em {tempo_str}"))

        except Exception as e:
            print(f"ERRO FATAL: {e}")
            self.master.after(0, lambda erro=str(e): messagebox.showerror("Erro", f"Ocorreu um erro:\n{erro}"))
            self.master.after(0, lambda: self.lbl_status.configure(text="Erro!"))
        finally:
            self.master.after(0, lambda: self.btn_run.configure(state="normal", text="PROCESSAR TUDO"))

    def _upd(self, v, t):
        p = v / t
        self.pb.set(p)
        self.lbl_status.configure(text=f"Excel: {v}/{t} ({int(p*100)}%)")