import os
import sys
import datetime
import tempfile
import shutil
import warnings
from pathlib import Path
import concurrent.futures
import multiprocessing

import streamlit as st

from service_access import gerar_mdb_access
from service_excel import extrair_info_template, processar_arquivo_isolado

# Suprimir avisos inofensivos
warnings.filterwarnings("ignore", message=".*missing ScriptRunContext.*")
warnings.filterwarnings("ignore", message=".*Process.*finalized.*")

# Configurar multiprocessing para Windows
if sys.platform == "win32":
    multiprocessing.set_start_method("spawn", force=True)

# --- FUNÇÃO PARA RESOURCE PATH (compatível com PyInstaller) ---
def resource_path(relative_path):
    """Retorna o caminho absoluto para o recurso, funciona para dev e PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="DDV - Migração",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- APLICAR ESTILO CUSTOMIZADO ---
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 1rem;
        color: #1f77b4;
    }
    .subtitle {
        font-size: 1.2rem;
        text-align: center;
        color: #666;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeeba;
        color: #856404;
    }
</style>
""", unsafe_allow_html=True)

# --- INICIALIZAR SESSION STATE ---
if "rotina_selecionada" not in st.session_state:
    st.session_state.rotina_selecionada = "SJ230133"
if "processando" not in st.session_state:
    st.session_state.processando = False
if "resultado_processamento" not in st.session_state:
    st.session_state.resultado_processamento = None

# --- SIDEBAR ---
st.sidebar.markdown("## ⚙️ Configuração")
rotinas_disponiveis = ["SJ230133", "SJ071930", "SJ071984"]
st.session_state.rotina_selecionada = st.sidebar.selectbox(
    "Selecione a Rotina:",
    options=rotinas_disponiveis,
    index=rotinas_disponiveis.index(st.session_state.rotina_selecionada)
)

st.sidebar.markdown("---")
st.sidebar.markdown("**ℹ️ Informações**")
st.sidebar.info(
    f"**Versão:** 2.0 (Streamlit)\n\n"
    f"**Rotina Atual:** {st.session_state.rotina_selecionada}\n\n"
    f"**Processamento:** Paralelo com CPU completa"
)

# --- CONTEÚDO PRINCIPAL ---
st.markdown('<div class="main-title">📊 DDV </div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Processamento de arquivos TXT para Access e Excel</div>', unsafe_allow_html=True)

# --- SEÇÃO 1: UPLOAD DE ARQUIVOS ---
st.markdown("## 📁 Etapa 1: Seleção de Arquivos")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### Header - F (TXT)")
    file_header = st.file_uploader(
        "Selecione o arquivo Header:",
        type=["txt"],
        key="header_file"
    )
    if file_header:
        st.success(f"✅ {file_header.name}")

with col2:
    st.markdown("### Detail - V (TXT)")
    file_detail = st.file_uploader(
        "Selecione o arquivo Detail:",
        type=["txt"],
        key="detail_file"
    )
    if file_detail:
        st.success(f"✅ {file_detail.name}")

st.markdown("---")

col3, col4 = st.columns(2)

with col3:
    st.markdown("### Índices de correção (TXT)")
    file_indices = st.file_uploader(
        "Selecione o arquivo de Índices (opcional):",
        type=["txt"],
        key="indices_file"
    )
    if file_indices:
        st.success(f"✅ {file_indices.name}")

with col4:
    st.markdown("### Diretório de Saída")
    output_dir = st.text_input(
        "Caminho para salvar os arquivos gerados:",
        value=str(Path.home() / "Desktop" / "DDV_Output"),
        key="output_dir"
    )

# --- SEÇÃO 2: PROCESSAMENTO ---
st.markdown("## ⚙️ Etapa 2: Processamento")

# Validação prévia
is_valid = file_header is not None and file_detail is not None
if not is_valid:
    st.warning("⚠️ Você precisa selecionar ao menos os arquivos Header e Detail para continuar.")

# Botão de processamento
col_btn1, col_btn2 = st.columns([2, 1])

with col_btn1:
    btn_processar = st.button(
        "🚀 PROCESSAR TUDO",
        type="primary",
        use_container_width=True,
        disabled=not is_valid
    )

# --- EXECUÇÃO DO PROCESSAMENTO ---
if btn_processar:
    st.session_state.processando = True
    
    try:
        tempo_inicio = datetime.datetime.now()
        
        # Criar diretório de saída
        os.makedirs(output_dir, exist_ok=True)
        
        # Criar arquivos temporários a partir dos uploads
        with tempfile.TemporaryDirectory() as temp_dir:
            # Salvar arquivos em temp
            temp_header = os.path.join(temp_dir, "header.txt")
            temp_detail = os.path.join(temp_dir, "detail.txt")
            temp_indices = os.path.join(temp_dir, "indices.txt")
            
            with open(temp_header, "wb") as f:
                f.write(file_header.getbuffer())
            
            with open(temp_detail, "wb") as f:
                f.write(file_detail.getbuffer())
            
            if file_indices:
                with open(temp_indices, "wb") as f:
                    f.write(file_indices.getbuffer())
            
            # Status container
            status_container = st.container()
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            with status_container:
                # --- CARREGAR ARQUIVOS ---
                status_text.info("📂 Lendo arquivos...")
                
                with open(temp_header, 'r', encoding='cp1252') as f:
                    lines_h = [l for l in f if len(l) >= 180]
                
                # Carregar índices
                idx_list = []
                if file_indices:
                    with open(temp_indices, 'r', encoding='cp1252') as f:
                        for linha in f:
                            if len(linha) >= 8:
                                try:
                                    dt = datetime.datetime.strptime(linha[:8], "%Y%m%d")
                                    val = float(linha[8:].replace(',', '.')) / 100.0
                                    idx_list.append((dt, val))
                                except (ValueError, IndexError):
                                    pass
                
                dt_lim = idx_list[-1][0] if idx_list else datetime.datetime.now()
                
                # Carregar detalhes
                with open(temp_detail, 'r', encoding='cp1252') as f:
                    linhas = f.readlines()
                
                lines_d_valid = [l for l in linhas if len(l) >= 120]
                
                from collections import defaultdict
                map_d = defaultdict(list)
                for l in lines_d_valid:
                    chave = l[:12].strip() + l[12:21].strip()
                    map_d[chave].append(l)
                
                # --- GERAR ACCESS ---
                status_text.info("🗄️ Gerando Access...")
                progress_bar.progress(10)
                
                p_tpl_mdb = resource_path(os.path.join("Templates", "MDB-Matriz.mdb"))
                p_tpl_xls = resource_path(os.path.join("Templates", "XLS-Matriz.xlsx"))
                
                # Validar templates
                if not os.path.exists(p_tpl_mdb):
                    raise FileNotFoundError(f"Template MDB não encontrado: {p_tpl_mdb}")
                if not os.path.exists(p_tpl_xls):
                    raise FileNotFoundError(f"Template XLS não encontrado: {p_tpl_xls}")
                
                ok_mdb, msg_mdb = gerar_mdb_access(
                    lines_h, 
                    lines_d_valid, 
                    output_dir, 
                    st.session_state.rotina_selecionada, 
                    p_tpl_mdb
                )
                
                if not ok_mdb:
                    st.warning(f"⚠️ {msg_mdb}")
                else:
                    st.success(f"✅ {msg_mdb}")
                
                # --- PROCESSAR EXCEL ---
                status_text.info("📊 Processando Excel...")
                progress_bar.progress(30)
                
                tpl_info = extrair_info_template(p_tpl_xls)
                
                # Construir tarefas
                tasks = [
                    (l, p_tpl_xls, output_dir, st.session_state.rotina_selecionada, 
                     idx_list, map_d.get(l[:12].strip() + l[12:21].strip(), []), 
                     dt_lim, tpl_info)
                    for l in lines_h
                ]
                
                tot = len(tasks)
                done = 0
                errs = []
                resultados = []
                
                # Usar ThreadPoolExecutor em vez de ProcessPoolExecutor (funciona melhor com Streamlit)
                with concurrent.futures.ThreadPoolExecutor(
                    max_workers=max(1, min(4, multiprocessing.cpu_count()-1))
                ) as exe:
                    futs = {exe.submit(processar_arquivo_isolado, t): t for t in tasks}
                    
                    for f in concurrent.futures.as_completed(futs):
                        try:
                            res = f.result()
                            done += 1
                            
                            if "ERRO" in res:
                                errs.append(res)
                            else:
                                resultados.append(res)
                            
                            progress = 30 + (int((done / tot) * 60))
                            progress_bar.progress(progress)
                            status_text.info(f"📊 Excel: {done}/{tot} ({int((done/tot)*100)}%)")
                        except Exception as e:
                            errs.append(f"ERRO ao processar: {str(e)}")
                            done += 1
                
                # --- RELATÓRIO FINAL ---
                tempo_fim = datetime.datetime.now()
                tempo_total = tempo_fim - tempo_inicio
                tempo_str = str(tempo_total).split('.')[0]
                
                progress_bar.progress(100)
                status_text.success("✅ Processamento concluído!")
                
                # Armazenar resultado
                st.session_state.resultado_processamento = {
                    'tempo_total': tempo_str,
                    'sucesso': tot - len(errs),
                    'erros': len(errs),
                    'detalhes_erros': errs[:10],  # Primeiros 10 erros
                    'output_dir': output_dir
                }
    
    except Exception as e:
        st.error(f"❌ Erro Durante Processamento: {str(e)}")
        st.session_state.resultado_processamento = None
    
    finally:
        st.session_state.processando = False

# --- MOSTRAR RESULTADO SE EXISTIR ---
if st.session_state.resultado_processamento:
    st.markdown("---")
    st.markdown("## 📋 Resultado do Processamento")
    
    res = st.session_state.resultado_processamento
    
    # Métricas
    col_met1, col_met2, col_met3, col_met4 = st.columns(4)
    
    with col_met1:
        st.metric("⏱️ Tempo Total", res['tempo_total'])
    
    with col_met2:
        st.metric("✅ Sucesso", res['sucesso'])
    
    with col_met3:
        st.metric("❌ Erros", res['erros'])
    
    with col_met4:
        st.metric("📁 Saída", res['output_dir'])
    
    st.markdown("---")
    
    # Detalhes de erros se existirem
    if res['erros'] > 0:
        with st.expander("🔍 Detalhes dos Erros"):
            for erro in res['detalhes_erros']:
                st.write(f"❌ {erro}")
    
    # Confirmação do sucesso
    if res['erros'] == 0:
        st.markdown(
            '<div class="success-box">🎉 Todos os arquivos foram processados com sucesso!</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div class="warning-box">⚠️ {res["erros"]} arquivo(s) apresentaram erros durante o processamento.</div>',
            unsafe_allow_html=True
        )

# --- RODAPÉ ---
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem; margin-top: 2rem;">
    <p>DDV - Migração de Dados | Versão 2.0 (Streamlit)</p>
</div>
""", unsafe_allow_html=True)
