import os
import datetime
import re
from copy import copy
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

def extrair_info_template(template_path):
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb["Receitas"]
    info = TemplateInfo()
    
    # 1. Linha Modelo (17) - Simplificado
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=17, column=col)
        info.estilos_linha_modelo[col] = {
            'number_format': cell.number_format,
            'font': copy(cell.font),
            'border': copy(cell.border),
            'alignment': copy(cell.alignment),
        }
        info.formulas_linha_modelo[col] = cell.value if _eh_formula(cell.value) else None

    # 2. Footer - Otimizado
    start_footer = 18
    for r in range(start_footer, ws.max_row + 20): 
        row_data = [
            {
                'value': (cell := ws.cell(row=r, column=c)).value,
                'number_format': cell.number_format,
                'font': copy(cell.font),
                'border': copy(cell.border),
                'alignment': copy(cell.alignment),
                'fill': copy(cell.fill)
            }
            for c in range(1, ws.max_column + 1)
        ]
        has_content = any(item['value'] for item in row_data)
        if has_content or r <= ws.max_row: 
            info.dados_footer.append(row_data)
        else: break 
    wb.close()
    return info

def processar_arquivo_isolado(args):
    (line_h, template_path, output_folder, rotina, indices, detalhes, dt_limite, tpl_info) = args
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

        wb = openpyxl.load_workbook(template_path)
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

        # Detalhes - Usa constantes congeladas (sem recriar)
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
            ws.cell(curr, 8, d['i'] or 0)  # Simplificado: 0 se falsy
            ws.cell(curr, 10, d['h'] or 0)

            for col in range(1, ws.max_column + 1):
                cell = ws.cell(curr, col)
                st = tpl_info.estilos_linha_modelo.get(col)
                fm = tpl_info.formulas_linha_modelo.get(col)
                
                # Aplicar estilo se existir
                if st:
                    cell.number_format = st['number_format']
                    cell.font = st['font']
                    cell.border = st['border']
                    cell.alignment = st['alignment']
                
                # Aplicar fórmula ou VLOOKUP padrão
                if fm:
                    cell.value = fm.replace("17", str(curr))
                elif col == 3:
                    cell.value = f"=VLOOKUP(A{curr},TOTINDICE!A:B,2,0)"
                    cell.number_format = '0.000000'
            curr += 1

        # Footer - Otimizado
        linha_fim_dados = curr - 1
        offset = curr - 18
        
        def processar_formula_footer(val, linha_fim, offset_val):
            """Processa fórmulas do footer em um único lugar."""
            if not _eh_formula(val):
                return val
            
            vu = val.upper()
            eh_soma = "SUM" in vu or "SOMA" in vu
            
            if eh_soma and ":" in vu:
                m = re.search(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", vu)
                if m and m.group(2) == "17":
                    return f"=SUM({m.group(1)}17:{m.group(3)}{linha_fim})"
            
            if not (eh_soma and "17" in val):
                def repl(m):
                    c_row = int(m.group(2))
                    return f"{m.group(1)}{c_row + offset_val}" if c_row >= 18 else m.group(0)
                val = re.sub(r"([A-Z]+)(\d+)", repl, val)
            
            return val
        
        for i, r_dat in enumerate(tpl_info.dados_footer):
            r_w = curr + i
            for j, c_dat in enumerate(r_dat):
                col = j + 1
                cell = ws.cell(r_w, col)
                val = processar_formula_footer(c_dat['value'], linha_fim_dados, offset)
                
                cell.value = val
                cell.number_format = c_dat['number_format']
                cell.font = c_dat['font']
                cell.border = c_dat['border']
                cell.alignment = c_dat['alignment']
                cell.fill = c_dat['fill']

        ws["C10"] = last_dt if last_dt else dt_limite
        ws["C10"].number_format = 'dd/mmm/yy'
        wb.save(path)
        wb.close()
        return nome_arq
    except Exception as e:
        return f"ERRO: {proc} - {str(e)}"