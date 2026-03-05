import os
import sys
import datetime
import tempfile
import warnings
import subprocess
import base64
from pathlib import Path
import concurrent.futures
import multiprocessing
from collections import defaultdict

import streamlit as st

from access import gerar_mdb_access
from excel import processar_arquivo_isolado

# Oculta avisos não críticos do Streamlit
warnings.filterwarnings("ignore", message=".*missing ScriptRunContext.*")
warnings.filterwarnings("ignore", message=".*Process.*finalized.*")

# Força o método 'spawn' no Windows para evitar travamentos de concorrência
if sys.platform == "win32":
    multiprocessing.set_start_method("spawn", force=True)

# --- Utilitários ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_base64_image(image_path):
    caminho_completo = resource_path(image_path)
    if not os.path.exists(caminho_completo):
        return ""
    try:
        with open(caminho_completo, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""

def abrir_explorador(caminho):
    try:
        if sys.platform == "win32":
            import ctypes
            caminho_norm = os.path.normpath(caminho)
            
            # Simula um toque na tecla ALT (0x12) para desativar o bloqueio de tela do Windows
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            
            # Abre o Explorer em primeiro plano
            ctypes.windll.shell32.ShellExecuteW(None, "explore", caminho_norm, None, None, 1)
            
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
        return True
    except Exception as e:
        import streamlit as st
        st.error(f"Falha ao abrir a pasta: {str(e)}")
        return False

def selecionar_pasta_windows(pasta_inicial=None):
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        
        if pasta_inicial is None or not os.path.exists(pasta_inicial):
            pasta_inicial = str(Path.home())
            
        pasta_selecionada = filedialog.askdirectory(
            initialdir=pasta_inicial,
            title="Selecione o diretório de saída"
        )
        
        root.destroy()
        
        if pasta_selecionada:
            return os.path.normpath(pasta_selecionada)
        return None
            
    except Exception as e:
        st.error(f"Erro de renderização do Tkinter: {str(e)}")
        return None

def ler_arquivo_com_fallback(caminho):
    """Tenta ler o arquivo como UTF-8; se falhar, usa ANSI (cp1252)."""
    for enc in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(caminho, 'r', encoding=enc) as f:
                return f.readlines()
        except (UnicodeDecodeError, LookupError):
            continue
            
    # Leitura de contingência ignorando caracteres corrompidos
    with open(caminho, 'r', encoding='cp1252', errors='replace') as f:
        return f.readlines()

# --- Callbacks de UI ---
def on_browse_click():
    pasta_selecionada = selecionar_pasta_windows(st.session_state.dir_saida)
    if pasta_selecionada:
        st.session_state.dir_saida = pasta_selecionada

def limpar_tudo():
    st.session_state.uploader_key += 1
    st.session_state.resultado_processamento = None
    st.session_state.dir_saida = ""
    st.session_state.processando = False

# --- Configuração Base de UI ---
st.set_page_config(
    page_title="DDV",
    page_icon=resource_path("icons/msaccess.jpg"),
    layout="wide",
    initial_sidebar_state="expanded"
)

b64_access = get_base64_image(os.path.join("icons", "msaccess.jpg"))
b64_file_open = get_base64_image(os.path.join("icons", "file_open.png"))

custom_css = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none !important;}
    .main-title { font-size: 2.5rem; font-weight: bold; text-align: center; margin-bottom: 1rem; color: #1f77b4; display: flex; align-items: center; justify-content: center; gap: 15px; }
    .subtitle { font-size: 1.2rem; text-align: center; color: #666; margin-bottom: 2rem; }
    .success-box, .error-box, .warning-box { padding: 1rem; border-radius: 0.5rem; border: 1px solid; }
    .success-box { background-color: #d4edda; border-color: #c3e6cb; color: #155724; }
    .error-box { background-color: #f8d7da; border-color: #f5c6cb; color: #721c24; }
    .warning-box { background-color: #fff3cd; border-color: #ffeeba; color: #856404; }
    [data-testid="stFileUploadDropzone"] > div > div > span, [data-testid="stFileUploaderDropzone"] > div > div > span { display: none !important; }
    [data-testid="stFileUploadDropzone"] > div > div > small, [data-testid="stFileUploaderDropzone"] > div > div > small { display: none !important; }
    [data-testid="stFileUploadDropzone"] > div > div::before, [data-testid="stFileUploaderDropzone"] > div > div::before { content: "Arraste e solte o arquivo aqui" !important; color: #31333F !important; font-size: 16px !important; font-weight: 500 !important; display: block !important; margin-bottom: 5px !important; }
"""

if b64_file_open:
    custom_css += f"""
    [data-testid="stFileUploadDropzone"] button p, [data-testid="stFileUploaderDropzone"] button p {{ display: none !important; }}
    [data-testid="stFileUploadDropzone"] button, [data-testid="stFileUploaderDropzone"] button {{ background-image: url("data:image/png;base64,{b64_file_open}") !important; background-size: 20px !important; background-position: center !important; background-repeat: no-repeat !important; color: transparent !important; width: 60px !important; margin: 0 auto !important; }}
    """
custom_css += "</style>"
st.markdown(custom_css, unsafe_allow_html=True)

# --- Gerenciamento de Estado da Sessão ---
if "rotina_selecionada" not in st.session_state: st.session_state.rotina_selecionada = "SJ230133"
if "processando" not in st.session_state: st.session_state.processando = False
if "resultado_processamento" not in st.session_state: st.session_state.resultado_processamento = None
if "dir_saida" not in st.session_state: st.session_state.dir_saida = ""
if "uploader_key" not in st.session_state: st.session_state.uploader_key = 0

# --- Renderização Sidebar ---
st.sidebar.markdown("## ⚙️ Configuração")
rotinas_disponiveis = ["SJ230133", "SJ071984"]
st.session_state.rotina_selecionada = st.sidebar.selectbox(
    "Selecione a Rotina:",
    options=rotinas_disponiveis,
    index=rotinas_disponiveis.index(st.session_state.rotina_selecionada)
)

# --- Renderização Principal ---
if b64_access:
    st.markdown(f'<div class="main-title"><img src="data:image/jpeg;base64,{b64_access}" width="40"> DDV </div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="main-title">📊 DDV </div>', unsafe_allow_html=True)

st.markdown('<div class="subtitle">Demonstrativo de Diferença de Vencimentos</div>', unsafe_allow_html=True)

st.markdown("## 📁 Etapa 1: Seleção de Arquivos")
col1, col2 = st.columns(2)

with col1:
    st.markdown("### Arquivo Header (TXT / TMP)")
    file_header = st.file_uploader("Selecione o arquivo Header:", key=f"header_file_{st.session_state.uploader_key}", type=["txt", "tmp"])
    if file_header: st.success(f"✅ {file_header.name}")

with col2:
    st.markdown("### Arquivo Detail (TXT / TMP)")
    file_detail = st.file_uploader("Selecione o arquivo Detail:", key=f"detail_file_{st.session_state.uploader_key}", type=["txt", "tmp"])
    if file_detail: st.success(f"✅ {file_detail.name}")

st.markdown("---")
col3, col4 = st.columns(2)

with col3:
    st.markdown("### Índices de Correção (TXT / TMP)")
    file_indices = st.file_uploader("Selecione o arquivo de Índices (opcional):", key=f"indices_file_{st.session_state.uploader_key}", type=["txt", "tmp"])
    if file_indices: st.success(f"✅ {file_indices.name}")

with col4:
    st.markdown("### 📁 Diretório de Saída")
    if file_header and file_detail:
        nome_base = Path(file_header.name).stem
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_saida = f"{nome_base}_{timestamp}"
        diretorio_base_desktop = str(Path.home() / "Desktop" / nome_saida)
        
        if not st.session_state.dir_saida:
            st.session_state.dir_saida = diretorio_base_desktop
            
        with st.container(border=True):
            st.markdown("<span style='color:#666; font-size:14px;'>Selecione o diretório para salvar (output):</span>", unsafe_allow_html=True)
            col_text, col_btn = st.columns([3, 1.2])
            with col_text:
                st.text_input("Caminho", key="dir_saida", label_visibility="collapsed")
            with col_btn:
                st.button("📁 Procurar", key="btn_browse_dir", on_click=on_browse_click, use_container_width=True)
    else:
        st.markdown("<div style='font-size: 14px; margin-bottom: 4px;'>&nbsp;</div>", unsafe_allow_html=True)
        st.warning("⚠️ Selecione os arquivos Header e Detail para definir o diretório de saída.")
        st.session_state.dir_saida = str(Path.home() / "Desktop" / "DDV_Output")

# --- Estrutura de Controles ---
st.markdown("## ⚙️ Etapa 2: Processamento")

is_valid = file_header is not None and file_detail is not None
if not is_valid:
    st.warning("⚠️ Você precisa selecionar ao menos os arquivos Header e Detail para continuar.")

col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])

with col_btn1:
    btn_processar = st.button("🚀 PROCESSAR TUDO", type="primary", use_container_width=True, disabled=not is_valid)
with col_btn2:
    btn_cancelar = st.button("🛑 CANCELAR", type="primary", use_container_width=True, disabled=not is_valid)
with col_btn3:
    btn_limpar = st.button("🧹 LIMPAR TUDO", type="secondary", use_container_width=True, on_click=limpar_tudo)

if btn_cancelar:
    st.warning("⚠️ Processamento cancelado pelo usuário.")
    st.session_state.resultado_processamento = None
    st.session_state.processando = False

# --- Core de Processamento ---
if btn_processar:
    st.session_state.processando = True
    try:
        tempo_inicio = datetime.datetime.now()
        hora_inicio_str = tempo_inicio.strftime("%H:%M:%S")
        
        diretorio_final = st.session_state.dir_saida
        os.makedirs(diretorio_final, exist_ok=True)
        
        st.info(f"▶️ **Processamento iniciado às:** `{hora_inicio_str}`")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_header = os.path.join(temp_dir, "header.txt")
            temp_detail = os.path.join(temp_dir, "detail.txt")
            temp_indices = os.path.join(temp_dir, "indices.txt")
            
            with open(temp_header, "wb") as f: f.write(file_header.getbuffer())
            with open(temp_detail, "wb") as f: f.write(file_detail.getbuffer())
            if file_indices:
                with open(temp_indices, "wb") as f: f.write(file_indices.getbuffer())
            
            # Container principal para os status
            status_container = st.container()
            
            with status_container:
                st.info("📂 Lendo e processando dados em memória...")

                lines_h_raw = ler_arquivo_com_fallback(temp_header)
                lines_h = [l for l in lines_h_raw if len(l) >= 180]
                
                idx_list = []
                if file_indices:
                    for linha in ler_arquivo_com_fallback(temp_indices):
                        if len(linha) >= 8:
                            try:
                                dt = datetime.datetime.strptime(linha[:8], "%Y%m%d")
                                val = float(linha[8:].replace(',', '.'))
                                idx_list.append((dt, val))
                            except (ValueError, IndexError):
                                pass
                
                dt_lim = idx_list[-1][0] if idx_list else datetime.datetime.now()
                
                lines_d_valid = []
                map_d = defaultdict(list)
                for l in ler_arquivo_com_fallback(temp_detail):
                    if len(l) >= 120:
                        lines_d_valid.append(l)
                        map_d[l[:12].strip() + l[12:21].strip()].append(l)
                
                # --- Progresso do Access ---
                st.markdown("#### 🗄️ Etapa A: Banco de Dados (Access)")
                access_status = st.empty()
                access_bar = st.progress(0)
                access_status.info("Gerando arquivo .mdb...")
                access_bar.progress(25) # Barra simulando o início do processo
                
                p_tpl_mdb = resource_path(os.path.join("Templates", "MDB-Matriz.mdb"))
                p_tpl_xls = resource_path(os.path.join("Templates", "XLS-Matriz.xlsx"))
                
                if not os.path.exists(p_tpl_mdb): raise FileNotFoundError(f"Template MDB ausente: {p_tpl_mdb}")
                if not os.path.exists(p_tpl_xls): raise FileNotFoundError(f"Template XLS ausente: {p_tpl_xls}")
                
                ok_mdb, msg_mdb = gerar_mdb_access(lines_h, lines_d_valid, diretorio_final, st.session_state.rotina_selecionada, p_tpl_mdb)
                
                access_bar.progress(100) # Enche a barra do Access ao finalizar
                if not ok_mdb: access_status.warning(f"⚠️ {msg_mdb}")
                else: access_status.success(f"✅ {msg_mdb}")
                
                # --- Progresso do Excel ---
                st.markdown("#### 📊 Etapa B: Planilhas Excel")
                excel_status = st.empty()
                excel_bar = st.progress(0)
                excel_status.info("Iniciando processamento paralelo...")
                
                tasks = [
                    (l, p_tpl_xls, diretorio_final, st.session_state.rotina_selecionada, 
                     idx_list, map_d.get(l[:12].strip() + l[12:21].strip(), []), dt_lim)
                    for l in lines_h
                ]
                
                tot = len(tasks)
                done = 0
                errs = []
                resultados = []
                
                if tot > 0:
                    # Execução em paralelo real contornando o GIL do Python
                    exe = concurrent.futures.ProcessPoolExecutor(max_workers=max(1, multiprocessing.cpu_count() - 1))
                    futs = {}
                    try:
                        futs = {exe.submit(processar_arquivo_isolado, t): t for t in tasks}
                        for f in concurrent.futures.as_completed(futs):
                            try:
                                res = f.result()
                                done += 1
                                if "ERRO" in res: errs.append(res)
                                else: resultados.append(res)
                                
                                pct = int((done / tot) * 100)
                                excel_bar.progress(pct) # Atualiza a barra de forma precisa (0 a 100)
                                excel_status.info(f"Processando: {done}/{tot} ({pct}%) planilhas concluídas")
                            except Exception as e:
                                errs.append(f"Falha na alocação do processo: {str(e)}")
                                done += 1
                    finally:
                        for f in futs: f.cancel()
                        exe.shutdown(wait=False)
                
                excel_bar.progress(100)
                excel_status.success("✅ Geração de planilhas finalizada!")
                
                tempo_fim = datetime.datetime.now()
                hora_fim_str = tempo_fim.strftime("%H:%M:%S")
                tempo_total = tempo_fim - tempo_inicio
                tempo_str = str(tempo_total).split('.')[0]
                
                st.session_state.resultado_processamento = {
                    'hora_inicio': hora_inicio_str,
                    'hora_fim': hora_fim_str,
                    'tempo_total': tempo_str,
                    'sucesso': tot - len(errs),
                    'erros': len(errs),
                    'detalhes_erros': errs[:10],
                    'output_dir': diretorio_final
                }
                
    except Exception as e:
        if type(e).__name__ in ('ScriptControlException', 'RerunException', 'StopException'):
            raise e
        st.error(f"❌ Erro Crítico de Execução: {str(e)}")
        st.session_state.resultado_processamento = None
    finally:
        st.session_state.processando = False

# --- Renderização de Resultados ---
if st.session_state.resultado_processamento:
    st.markdown("---")
    st.markdown("## 📋 Resumo do Processamento")
    res = st.session_state.resultado_processamento
    
    col_met1, col_met2, col_met3, col_met4 = st.columns(4)
    with col_met1: st.metric("🕐 Início", res['hora_inicio'])
    with col_met2: st.metric("🕑 Término", res['hora_fim'])
    with col_met3: st.metric("⏱️ Duração", res['tempo_total'])
    with col_met4: st.metric("📊 Arquivos Gerados", f"{res['sucesso']} ✅ | {res['erros']} ❌")
    
    st.markdown("---")
    col_label, col_btn_result = st.columns([3, 1])
    with col_label:
        st.markdown("### 📁 Localização dos Arquivos")
        st.caption(res['output_dir'])
    with col_btn_result:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📁 Abrir Pasta", key="btn_browse_result", type="primary", use_container_width=True):
            abrir_explorador(res['output_dir'])
            st.success("✅ Acesso concedido via Windows Explorer.")
    
    st.markdown("---")
    if res['erros'] > 0:
        with st.expander("🔍 Rastreamento de Erros"):
            for erro in res['detalhes_erros']: st.write(f"❌ {erro}")
    
    if res['erros'] == 0:
        st.markdown('<div class="success-box">🎉 Execução finalizada sem advertências.</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="warning-box">⚠️ Sub-processos corrompidos detectados: {res["erros"]} arquivo(s).</div>', unsafe_allow_html=True)

st.markdown("---")