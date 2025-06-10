import os
import time
from io import StringIO
from tkinter.filedialog import askdirectory

import language_tool_python as lt
from openpyxl import Workbook

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


def extract_text(path):
    output_string = StringIO()
    with open(path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
    return output_string.getvalue()


def analyze_text(text):
    tool = lt.LanguageTool("pt-BR", config={'requestLimit': 10, 'timeoutRequestLimit': 60})
    matches = tool.check(text)
    return len(matches)


def main():
    # Obter caminhos de entrada e saída
    path = askdirectory(title='Caminho Data salva')
    path_resultado = askdirectory(title='Caminho onde o resultado da pesquisa deve ser salvo')
    path_saida_original = askdirectory(title='Caminho de Saida')
    path_texto = askdirectory(title='Caminho Texto')

    # Criar planilha de resultados
    wb = Workbook()
    ws = wb.active
    ws.title = "Erros Ortográficos"

    # Percorrer arquivos PDF
    files = []
    for root, _, filenames in os.walk(path):
        for filename in filenames:
            if filename.endswith('.pdf'):
                files.append(os.path.join(root, filename))

    for i, pdf_file in enumerate(files):
        print(f"Analisando arquivo {i + 1} de {len(files)}: {pdf_file}")
        try:
            # Extrair texto do PDF e salvar em arquivo de texto
            text = extract_text(pdf_file)
            txt_file = pdf_file.replace('.pdf', '.txt')
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write(text)

            # Analisar texto e contar erros
            num_errors = analyze_text(text)

            # Adicionar informações na planilha
            filename = os.path.basename(pdf_file)
            path_parts = pdf_file.split(os.path.sep)
            company = path_parts[-2]
            year = path_parts[-1].split()[0]
            ws.append([company, year, filename, num_errors])

        except Exception as e:
            print(f"Erro ao analisar arquivo {pdf_file}: {str(e)}")

    # Salvar planilha
    wb.save(os.path.join(path_saida_original, 'Erros Ortográficos.xlsx'))


if __name__ == '__main__':
    start_time = time.time()
    main()
    elapsed_time = time.time() - start_time
   

