import os
import re
import time
import language_tool_python as lt
import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog

# Função para analisar o texto com a biblioteca language_tool_python
def analyze_text(text):
    tool = lt.LanguageTool("pt-BR")
    return tool.check(text)

# Função para analisar PDFs
def analisar_pdf(arquivo_pdf):
    # Extrair texto do PDF
    texto = Path(arquivo_pdf).read_text(encoding='latin1')

    # Remover quebras de página e caracteres especiais
    texto = re.sub(r"\n|\r|\t|-", " ", texto)

    # Verificar a quantidade de páginas do PDF
    qtde_paginas = texto.count('\x0c') + 1
    print(f"\tAnalisando relatório {arquivo_pdf} ({qtde_paginas} páginas)")

    # Analisar o texto com a biblioteca language_tool_python
    erros = analyze_text(texto)

    # Criar um dataframe para salvar os erros encontrados
    df_erros = pd.DataFrame(columns=["Arquivo", "Erro", "Detalhes", "Sugestão de correção"])

    # Preencher o dataframe com os erros encontrados
    for erro in erros:
        new_row = pd.DataFrame({"Arquivo": arquivo_pdf,"Erro": erro.ruleId,"Detalhes": erro.message,"Sugestão de correção": erro.replacements[0] if erro.replacements else ""}, index=[0])
        df_erros = pd.concat([df_erros, new_row], ignore_index=True)

    return df_erros

# Abrir janela de seleção de pasta com o Tkinter
janela = Tk()
janela.withdraw()
pasta = filedialog.askdirectory()

# Criar um dataframe vazio para salvar os erros de todas as empresas
df_erros_total = pd.DataFrame(columns=["Arquivo", "Erro", "Detalhes", "Sugestão de correção"])

# Exemplo de uso da função para analisar todos os arquivos PDF em uma pasta
for empresa in os.listdir(pasta):
    caminho_empresa = os.path.join(pasta, empresa)
    if os.path.isdir(caminho_empresa):
        print(f"Analisando empresa {empresa}...")
        df_erros_empresa = pd.DataFrame(columns=["Arquivo", "Erro", "Detalhes", "Sugestão de correção"])
        for arquivo in os.listdir(caminho_empresa):
            if arquivo.endswith(".pdf"):
                caminho_arquivo = os.path.join(caminho_empresa, arquivo)
                df_erros_arquivo = analisar_pdf(caminho_arquivo)
                df_erros_empresa = pd.concat([df_erros_empresa, df_erros_arquivo], ignore_index=True)
        # Salvar os erros da empresa em um arquivo CSV
        nome_arquivo_saida = os.path.join(caminho_empresa, f"{empresa}_erros.xlsx")
        df_erros_empresa.to_excel(nome_arquivo_saida, index=False)
        # Adicionar os erros da empresa ao dataframe de erros total
        df_erros_total = df_erros_total.append(df_erros_empresa, ignore_index=True)

# Salvar todos os erros em um arquivo Excel
nome_arquivo_saida_total = os.path.join(pasta, "erros_total.xlsx")
df_erros_total.to_excel(nome_arquivo_saida_total, index=False)

print("Processo concluído com sucesso!")

