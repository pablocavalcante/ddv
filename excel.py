import os
import datetime
import re
from itertools import groupby
import openpyxl
 
# Constantes globais - evita recriar em cada execução
CODIGOS_IPREM = frozenset({"5001", "6013", "6017", "7012", "PREV", "RPPS"})
CODIGOS_HSPM = frozenset({"5101", "6015", "7011", "HSPM"})
 
class TemplateInfo:
    def __init__(self):
        self.formulas_linha_modelo = {} 
        self.estilos_linha_modelo = {}  
        self.dados_footer = []          
 
def _eh_formula(valor):
    """Verifica se o valor é uma fórmula."""
    return isinstance(valor, str) and "=" in valor
<<<<<<< HEAD
 
=======

>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
def extrair_info_template(wb):
    """
    Extrai as informações de estilo baseadas no ID interno do openpyxl (_style).
    Recebe o workbook (wb) já aberto para não ler o disco duas vezes.
    """
    ws = wb["Receitas"]
    info = TemplateInfo()
    
    # 1. Linha Modelo (17) - Extremamente Simplificado e Rápido
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=17, column=col)
        # OTIMIZAÇÃO: Copia apenas o ID numérico do estilo. Evita cópias pesadas de memória.
        info.estilos_linha_modelo[col] = cell._style
        info.formulas_linha_modelo[col] = cell.value if _eh_formula(cell.value) else None
 
    # 2. Footer - Otimizado
    start_footer = 18
    for r in range(start_footer, ws.max_row + 20): 
        row_data = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            row_data.append({
                'value': cell.value,
                '_style': cell._style # OTIMIZAÇÃO
            })
            
        has_content = any(item['value'] for item in row_data)
        if has_content or r <= ws.max_row: 
            info.dados_footer.append(row_data)
        else: 
            break 
            
    return info
 
def processar_arquivo_isolado(args):
    # OTIMIZAÇÃO: Não recebe o tpl_info via argumento. Acabaram-se os gargalos do Pickling!
    (line_h, template_path, output_folder, rotina, indices, detalhes, dt_limite) = args
    try:
        proc = line_h[0:12].strip()
        rf = line_h[12:21].strip()
        autor = line_h[87:119].strip()
        padrao = line_h[171:180].strip()
 
        nome_arq = f"{rotina}_{proc}-{rf}.xlsx"
        day_folder = os.path.join(output_folder, datetime.datetime.now().strftime("%Y-%m-%d"))
        p_folder = os.path.join(day_folder, f"Processo_{proc} - CD 1")
        os.makedirs(p_folder, exist_ok=True)
        path = os.path.join(p_folder, nome_arq)
<<<<<<< HEAD
 
=======

>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
        # O trabalhador (processo) abre o ficheiro uma única vez de forma independente
        wb = openpyxl.load_workbook(template_path)
        
        # Lê a informação do template usando o ficheiro que acabou de abrir
        tpl_info = extrair_info_template(wb)
        
        ws = wb["Receitas"]
        ws_idx = wb["TOTINDICE"]
 
        # Indices
        r_idx = ws_idx.max_row + 1 if ws_idx.max_row > 1 else 2
        existing = {r[0] for r in ws_idx.iter_rows(min_col=1, max_col=1, values_only=True) if isinstance(r[0], datetime.datetime)}
        
        for dt, val in indices:
            if dt not in existing:
                ws_idx.cell(r_idx, 1, dt).number_format = 'dd/mm/yyyy'
                ws_idx.cell(r_idx, 2, val).number_format = '0.000000'
                r_idx += 1
 
        ws["C7"] = proc
        ws["C8"] = f"{autor} - RF: {rf}"
<<<<<<< HEAD
 
=======

>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
        # Detalhes - Usa constantes congeladas
        def sort_key(x): return x[23:27] + x[21:23]
        detalhes.sort(key=sort_key)
        
        rows_data = []
        last_dt = None
 
        for key, group in groupby(detalhes, key=sort_key):
            g = list(group)
            t_venc = sum(float(l[86:96]) for l in g) / 100.0
            t_desc = sum(float(l[116:126]) for l in g) / 100.0
            v_iprem = sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_IPREM) / 100.0
            v_hspm = sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_HSPM) / 100.0
            
            val_q = max(0, (t_venc - t_desc) + v_iprem + v_hspm) if padrao.startswith("PR") else (t_venc - t_desc) + v_iprem + v_hspm
 
            ano, mes = int(key[:4]), int(key[4:])
            nxt = datetime.datetime(ano, mes, 28) + datetime.timedelta(days=4)
            last_dt = nxt - datetime.timedelta(days=nxt.day)
 
            rows_data.append({'dt': last_dt, 'q': val_q, 'i': v_iprem, 'h': v_hspm})
 
        curr = 17
        for d in rows_data:
            ws.cell(curr, 1, d['dt'])
            ws.cell(curr, 2, d['q'])
<<<<<<< HEAD
=======
            ws.cell(curr, 5, 0)
>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
            ws.cell(curr, 8, d['i'] or 0)  
            ws.cell(curr, 10, d['h'] or 0)
 
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(curr, col)
                st_id = tpl_info.estilos_linha_modelo.get(col)
                fm = tpl_info.formulas_linha_modelo.get(col)
                
                # APLICAÇÃO DE ESTILO RELÂMPAGO
                if st_id is not None:
                    cell._style = st_id
                
                # Aplicar fórmula ou VLOOKUP padrão
                if fm:
                    cell.value = fm.replace("17", str(curr))
                elif col == 3:
<<<<<<< HEAD
                    cell.value = f"=VLOOKUP(A{curr},TOTINDICE!A:B,2,0)"
                    cell.number_format = '0.000000' # Garante formatação apenas na formula dinâmica
            curr += 1
 
=======
                    cell.value = f'=IFERROR(VLOOKUP(A{curr},TOTINDICE!A:B,2,0), "")'
                    cell.number_format = '0.000000' # Garante formatação apenas na formula dinâmica
            curr += 1

>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
        # Footer
        linha_fim_dados = curr - 1
        offset = curr - 18
        
        def processar_formula_footer(val, linha_fim, offset_val):
            """Atualiza as fórmulas do rodapé usando Cláusulas de Guarda (Early Return)."""
            
            # 1. Cláusula de Guarda: Se não for fórmula, devolve como está e encerra.
            if not _eh_formula(val):
                return val
                
            vu = val.upper()
            
            # 2. Tratamento Direto: Somas dos totais (Ex: =SUM(B17:B20))
            if ("SUM" in vu or "SOMA" in vu) and ":" in vu:
                m = re.search(r"([A-Z]+)17:([A-Z]+)(\d+)", vu)
                if m:
                    # Achou a soma principal? Reescreve e já sai da função!
                    return f"=SUM({m.group(1)}17:{m.group(2)}{linha_fim})"

            # 3. Caso Geral: Tratamento das outras fórmulas da margem/rodapé
            def deslocar_linha(m):
                linha_atual = int(m.group(2))
                return f"{m.group(1)}{linha_atual + offset_val}" if linha_atual >= 18 else m.group(0)
                
            # Atualiza os números das linhas para empurrar o rodapé para baixo
            nova_formula = re.sub(r"([A-Z]+)(\d+)", deslocar_linha, val)
            
            # 4. Escudo Final: Aplica o IFERROR caso ainda não tenha
            if "IFERROR" not in vu and "SEERRO" not in vu:
                return f'=IFERROR({nova_formula[1:]}, "")'
                
            return nova_formula
        
        for i, r_dat in enumerate(tpl_info.dados_footer):
            r_w = curr + i
            for j, c_dat in enumerate(r_dat):
                col = j + 1
                cell = ws.cell(r_w, col)
                val = processar_formula_footer(c_dat['value'], linha_fim_dados, offset)
                
                cell.value = val
                # APLICAÇÃO DE ESTILO RELÂMPAGO NO FOOTER
                if c_dat['_style'] is not None:
                    cell._style = c_dat['_style']
<<<<<<< HEAD
 
=======

>>>>>>> 3c38a46a8546e8eac0e320fc4129457d3cb0632f
        ws["C10"] = last_dt if last_dt else dt_limite
        ws["C10"].number_format = 'dd/mmm/yy'
        wb.save(path)
        wb.close()
        return nome_arq
        
    except Exception as e:
        return f"ERRO: {proc} - {str(e)}"