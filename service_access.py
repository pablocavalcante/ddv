import os
import shutil
import datetime
import pyodbc

def _converter_para_float(val_str):
    """Converte string para float com tratamento de erro."""
    try:
        return float(val_str) / 100.0
    except (ValueError, TypeError):
        return 0.0

def gerar_mdb_access(header_lines, detail_lines_raw, output_folder, rotina, template_mdb_path):
    """
    Popula o MDB Matriz com os dados lidos do TXT.
    """
    nome_mdb = f"{rotina}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.mdb"
    caminho_final = os.path.join(output_folder, nome_mdb)
    
    try:
        # Validações iniciais (fail-fast)
        if not os.path.exists(template_mdb_path):
            return False, f"Template MDB não encontrado: {template_mdb_path}"
        
        shutil.copy(template_mdb_path, caminho_final)
        
        # Buscar driver ODBC do Access
        drivers = [x for x in pyodbc.drivers() if 'Access' in x]
        if not drivers:
            return False, "Driver ODBC do Access não encontrado no Windows."
        
        driver_name = drivers[0] 
        conn_str = fr'DRIVER={{{driver_name}}};DBQ={caminho_final};'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        # Preparar dados HEADER usando comprehension
        dados_header = [
            (
                line[0:12],      # Processo
                line[12:21],     # RF
                line[23:27],     # Data Ref Ano
                line[21:23],     # Data Ref Mes
                line[27:37],     # Data Processamento
                line[37:87],     # Observacao
                line[87:119],    # Autor
                line[119:171],   # Cargo
                line[171:180],   # Padrao
                line[180:182],   # Qtde Dias
                line[182:203]    # Auto
            )
            for line in header_lines if len(line) >= 180
        ]
        
        # Preparar dados DETAIL usando comprehension
        dados_detail = [
            (
                line[0:12],                        # Processo
                line[12:21],                       # RF
                line[23:27],                       # Data Ref Ano
                line[21:23],                       # Data Ref Mes
                line[27:31],                       # Codigo
                line[31:66],                       # Significado
                _converter_para_float(line[66:76]),  # Recebido
                _converter_para_float(line[76:86]),  # A Receber
                _converter_para_float(line[86:96]),  # Dif Venc
                _converter_para_float(line[96:106]), # Descontado
                _converter_para_float(line[106:116]),# A Descontar
                _converter_para_float(line[116:126]) # Dif Desc
            )
            for line in detail_lines_raw if len(line) >= 120
        ]
        
        # Inserção em lote - apenas se houver dados
        if dados_header:
            sql_h = """
                INSERT INTO Header (
                    Processo, RF, Data_Ref_Ano, Data_Ref_Mes, Data_Processamento, 
                    Observacao, Autor, Cargo, Padrao, Qtde_Dias, Auto
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.executemany(sql_h, dados_header)

        if dados_detail:
            sql_d = """
                INSERT INTO Detail (
                    Processo, RF, Data_Ref_Ano, Data_Ref_Mes, Codigo, 
                    Significado, Recebido, A_Receber, Dif_Venc, 
                    Descontado, A_Descontar, Dif_Desc
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.executemany(sql_d, dados_detail)

        conn.commit()
        conn.close()
        
        return True, f"Sucesso! Gerado: {nome_mdb}"

    except Exception as e:
        return False, f"Erro Access: {str(e)}"