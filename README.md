# DDV - Demonstrativo de Diferença de Vencimentos

## 🎯 Para que serve?

O **DDV** é a modernização de um sistema originalmente desenvolvido em VB6. Seu objetivo é realizar a leitura automatizada de arquivos `.txt` (extraídos do Mainframe) e gerar planilhas Excel formatadas, além de um banco de dados Access (`.mdb`) para facilitar a visualização das informações.

---

## 🚀 Processamento dos Arquivos

Para rodar o sistema e gerar os seus relatórios, siga o passo a passo abaixo:

1. **Baixe os arquivos de origem:** Os dados base para cada processo são exportados do `MAINFRAME`. Acesse o **FileZilla** e faça o download dos arquivos correspondentes ao processo:
<div align="center">
  <img width="175" height="89" alt="Image" src="https://github.com/user-attachments/assets/a36ae5e6-5e60-488a-a75d-27d906521ec0" />
</div>

2. **Insira os dados no sistema:** Os arquivos `.txt` possuem estruturas específicas. Na interface web do DDV, clique ou arraste os arquivos para os seus respectivos campos:
   - **Arquivo Header:** Insira o arquivo terminado em **_F** (`.txt`)
   - **Arquivo Detail:** Insira o arquivo terminado em **_V** (`.txt`)
   - **Índices de Correção:** Insira o arquivo contendo os índices (`.txt`)

3. **Escolha o destino:** No campo "Diretório de Saída", clique em "Procurar" e selecione a pasta do seu computador onde você deseja que o sistema salve os arquivos finais (Excel e Access) gerados.
<img width="1915" height="946" alt="Image" src="https://github.com/user-attachments/assets/83c0063c-93e0-4fca-9232-c426e29667f2" />

4. **Inicie o cálculo:** Com tudo preenchido, clique no botão azul **"🚀 PROCESSAR TUDO"**. 
<img width="1619" height="990" alt="Image" src="https://github.com/user-attachments/assets/8bc33838-eed3-4396-b54f-5c0f19c0e896" />

5. **Acesse seus resultados:** Aguarde a barra de progresso terminar. Ao final, o sistema mostrará uma mensagem de sucesso. Basta clicar no botão **"📁 Abrir Pasta"** para visualizar o Access e as suas planilhas prontas!
