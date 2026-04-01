nome_do_arquivo = "RET_ACAO_V.txt"

# Estes nós já sabemos o que são, então vamos ignorá-los para limpar a tela
codigos_conhecidos = {'6015', '6017', '7011', '7012', '5001', '5101'}
dicionario_codigos = {}

print("\nLendo o arquivo e traduzindo todos os códigos...")

try:
    with open(nome_do_arquivo, "r", encoding="latin-1") as f:
        for linha in f:
            if len(linha) > 70:
                codigo = linha[27:31].strip()
                
                # Se for um código válido, que não conhecemos e ainda não guardamos
                if codigo and codigo not in codigos_conhecidos and codigo not in dicionario_codigos:
                    nome_do_desconto = linha[31:70].strip()
                    dicionario_codigos[codigo] = nome_do_desconto

    print("\n=== TODOS OS CÓDIGOS DESCONHECIDOS ===")
    for cod in sorted(dicionario_codigos.keys()):
        print(f"[{cod}] - {dicionario_codigos[cod]}")
    print("======================================\n")

except FileNotFoundError:
    print("Arquivo Detail.txt não encontrado.")