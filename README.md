# DDV - Demonstrativo de Diferença de Vencimentos

## 🎯 Para que serve?

O **DDV** é a modernização de um sistema originalmente desenvolvido em VB6. Seu objetivo é realizar a leitura automatizada de arquivos `.txt` (extraídos do Mainframe) e gerar planilhas Excel formatadas, além de um banco de dados Access (`.mdb`) para facilitar a visualização das informações.

---

## 🚀 Processamento dos Arquivos

Para rodar o sistema e gerar os seus relatórios, siga o passo a passo abaixo:

1. **Baixe os arquivos de origem:** Os dados base para cada processo são exportados do `MAINFRAME`. Acesse o **FileZilla** e faça o download dos arquivos correspondentes ao processo:<br><br>
<div align="center">
  <img width="175" height="89" alt="Image" src="https://github.com/user-attachments/assets/a36ae5e6-5e60-488a-a75d-27d906521ec0" />
</div><br><br>

2. **Insira os dados no sistema:** Os arquivos `.txt` possuem estruturas específicas. Na interface web do DDV, clique ou arraste os arquivos para os seus respectivos campos:<br><br>

   - **Arquivo Header:** Insira o arquivo terminado em **_F** (`.txt`)
   - **Arquivo Detail:** Insira o arquivo terminado em **_V** (`.txt`)
   - **Índices de Correção:** Insira o arquivo contendo os índices (`.txt`)<br><br>

3. **Escolha o destino:** No campo "Diretório de Saída", clique em "Procurar" e selecione a pasta do seu computador onde você deseja que o sistema salve os arquivos finais (Excel e Access) gerados.<br><br>
<img width="1915" height="946" alt="Image" src="https://github.com/user-attachments/assets/83c0063c-93e0-4fca-9232-c426e29667f2" /><br><br>

4. **Inicie o cálculo:** Com tudo preenchido, clique no botão azul **"🚀 PROCESSAR TUDO"**. <br><br>
<img width="1106" height="623" alt="Image" src="https://github.com/user-attachments/assets/aaa74c53-f80d-4309-b10a-91a63bef99b7" /><br><br>

5. **Acesse seus resultados:** Aguarde a barra de progresso terminar. Ao final, o sistema mostrará uma mensagem de sucesso. Basta clicar no botão **"📁 Abrir Pasta"** para visualizar o Access e as suas planilhas prontas!
<img width="1084" height="551" alt="Image" src="https://github.com/user-attachments/assets/f166cff2-69fc-4dad-9334-641024367693" />
