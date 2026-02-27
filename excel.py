import os
import datetime
import re
from itertools import groupby
import openpyxl

# Constantes globais
CODIGOS_IPREM   = frozenset({"5001", "6013", "6017", "7012", "PREV", "RPPS"})
CODIGOS_HSPM    = frozenset({"5101", "6015", "7011", "HSPM"})

# Preencha com os códigos corretos do seu sistema para FUNFIN e FUNPREV
CODIGOS_FUNFIN  = frozenset()   # ex: {"XXXX", "YYYY"}
CODIGOS_FUNPREV = frozenset()   # ex: {"ZZZZ", "WWWW"}

# Mapeamento de colunas do novo template (XLS-MATRIZ.xlsx)
# A=1  B=2  C=3  D=4  E=5  F=6  G=7
# dd/mmm/aa | Quantum | Quan.Debeat.atualizado | IPREM | HSPM | TOT_FUNFIN | TOT_FUNPREV
COL_DATA     = 1
COL_QUANTUM  = 2
COL_ATUALIZ  = 3   # fórmula =B*$D$10/VLOOKUP(...)
COL_IPREM    = 4   # era col 8 no template antigo
COL_HSPM     = 5   # era col 10 no template antigo
COL_FUNFIN   = 6   # nova coluna
COL_FUNPREV  = 7   # nova coluna

# Linha modelo e início do footer no novo template
LINHA_MODELO  = 17
START_FOOTER  = 20   # era 18 no template antigo


class TemplateInfo:
    def __init__(self):
        self.estilos_linha_modelo  = {}
        self.formulas_linha_modelo = {}
        self.dados_footer          = []


def _eh_formula(valor):
    return isinstance(valor, str) and valor.startswith("=")


def extrair_info_template(wb):
    ws   = wb["Receitas"]
    info = TemplateInfo()

    # 1. Estilos da linha modelo (17)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=LINHA_MODELO, column=col)
        info.estilos_linha_modelo[col]  = cell._style
        # Ignora a fórmula quebrada da col C — será sempre reconstruída no código
        info.formulas_linha_modelo[col] = cell.value if (_eh_formula(cell.value) and col != COL_ATUALIZ) else None

    # 2. Footer a partir da linha 20
    for r in range(START_FOOTER, ws.max_row + 20):
        row_data = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            row_data.append({'value': cell.value, '_style': cell._style})

        has_content = any(item['value'] for item in row_data)
        if has_content or r <= ws.max_row:
            info.dados_footer.append(row_data)
        else:
            break

    return info


def processar_arquivo_isolado(args):
    (line_h, template_path, output_folder, rotina, indices, detalhes, dt_limite) = args
    try:
        proc   = line_h[0:12].strip()
        rf     = line_h[12:21].strip()
        autor  = line_h[87:119].strip()
        padrao = line_h[171:180].strip()

        nome_arq   = f"{rotina}_{proc}-{rf}.xlsx"
        day_folder = os.path.join(output_folder, datetime.datetime.now().strftime("%Y-%m-%d"))
        p_folder   = os.path.join(day_folder, f"Processo_{proc} - CD 1")
        os.makedirs(p_folder, exist_ok=True)
        path = os.path.join(p_folder, nome_arq)

        wb       = openpyxl.load_workbook(template_path)
        tpl_info = extrair_info_template(wb)
        ws       = wb["Receitas"]
        ws_idx   = wb["TOTINDICE"]

        # ── Índices ────────────────────────────────────────────────────────────
        r_idx    = ws_idx.max_row + 1 if ws_idx.max_row > 1 else 2
        existing = {
            r[0] for r in ws_idx.iter_rows(min_col=1, max_col=1, values_only=True)
            if isinstance(r[0], datetime.datetime)
        }
        for dt, val in indices:
            if dt not in existing:
                ws_idx.cell(r_idx, 1, dt).number_format = 'dd/mm/yyyy'
                ws_idx.cell(r_idx, 2, val).number_format = '0.000000'
                r_idx += 1

        # ── Cabeçalho ──────────────────────────────────────────────────────────
        ws["C7"] = proc
        ws["C8"] = f"{autor} - RF: {rf}"

        # ── Detalhes ────────────────────────────────────────────────────────────
        def sort_key(x): return x[23:27] + x[21:23]
        detalhes.sort(key=sort_key)

        rows_data = []
        last_dt   = None

        for key, group in groupby(detalhes, key=sort_key):
            g = list(group)

            t_venc   = sum(float(l[86:96])   for l in g) / 100.0
            t_desc   = sum(float(l[116:126]) for l in g) / 100.0

            v_iprem  = sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_IPREM)  / 100.0
            v_hspm   = sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_HSPM)   / 100.0
            v_funfin = sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_FUNFIN) / 100.0
            v_funprev= sum(float(l[116:126]) for l in g if l[27:31] in CODIGOS_FUNPREV)/ 100.0

            val_q = (
                max(0, (t_venc - t_desc) + v_iprem + v_hspm)
                if padrao.startswith("PR")
                else (t_venc - t_desc) + v_iprem + v_hspm
            )

            ano, mes = int(key[:4]), int(key[4:])
            nxt     = datetime.datetime(ano, mes, 28) + datetime.timedelta(days=4)
            last_dt = nxt - datetime.timedelta(days=nxt.day)

            rows_data.append({
                'dt': last_dt,
                'q' : val_q,
                'i' : v_iprem,
                'h' : v_hspm,
                'f' : v_funfin,
                'p' : v_funprev,
            })

        # ── Escrita das linhas ──────────────────────────────────────────────────
        curr = LINHA_MODELO
        for d in rows_data:
            # 1. Grava os valores nas colunas corretas ANTES de aplicar estilos
            ws.cell(curr, COL_DATA,    d['dt'])
            ws.cell(curr, COL_QUANTUM, d['q'])
            ws.cell(curr, COL_IPREM,   d['i'] or 0)
            ws.cell(curr, COL_HSPM,    d['h'] or 0)
            ws.cell(curr, COL_FUNFIN,  d['f'] or 0)
            ws.cell(curr, COL_FUNPREV, d['p'] or 0)

            # 2. Aplica estilos e fórmulas coluna a coluna
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(curr, col)

                # Estilo da linha-modelo
                st_id = tpl_info.estilos_linha_modelo.get(col)
                if st_id is not None:
                    cell._style = st_id

                # Coluna C: sempre reconstrói a fórmula correta (sem #REF!)
                if col == COL_ATUALIZ:
                    cell.value = (
                        f"=B{curr}*$D$10"
                        f"/VLOOKUP(A{curr},TOTINDICE!$A:$B,2,0)"
                    )
                    continue

                # Demais fórmulas da linha-modelo (ajusta nº de linha)
                fm = tpl_info.formulas_linha_modelo.get(col)
                if fm:
                    cell.value = fm.replace(str(LINHA_MODELO), str(curr))

            curr += 1

        # ── Footer ──────────────────────────────────────────────────────────────
        linha_fim_dados = curr - 1
        offset          = curr - START_FOOTER

        def processar_formula_footer(val, linha_fim, offset_val):
            if not _eh_formula(val):
                return val

            vu      = val.upper()
            eh_soma = "SUM" in vu or "SOMA" in vu

            # SUM com intervalo iniciando em LINHA_MODELO → expande até o fim dos dados
            if eh_soma and ":" in vu:
                m = re.search(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", vu)
                if m and int(m.group(2)) == LINHA_MODELO:
                    return f"=SUM({m.group(1)}{LINHA_MODELO}:{m.group(3)}{linha_fim})"

            # Demais fórmulas: desloca referências de linhas >= START_FOOTER
            if not (eh_soma and str(LINHA_MODELO) in val):
                def repl(m):
                    c_row = int(m.group(2))
                    return (
                        f"{m.group(1)}{c_row + offset_val}"
                        if c_row >= START_FOOTER
                        else m.group(0)
                    )
                val = re.sub(r"([A-Z]+)(\d+)", repl, val)

            return val

        for i, r_dat in enumerate(tpl_info.dados_footer):
            r_w = curr + i
            for j, c_dat in enumerate(r_dat):
                col  = j + 1
                cell = ws.cell(r_w, col)
                val  = processar_formula_footer(c_dat['value'], linha_fim_dados, offset)
                cell.value = val
                if c_dat['_style'] is not None:
                    cell._style = c_dat['_style']

        # ── Data de atualização ─────────────────────────────────────────────────
        ws["C10"] = last_dt if last_dt else dt_limite
        ws["C10"].number_format = 'dd/mmm/yy'

        wb.save(path)
        wb.close()
        return nome_arq

    except Exception as e:
        return f"ERRO: {proc} - {str(e)}"