from io import StringIO

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

import language_tool_python as lt

import os
#from tkinter.filedialog import askdirectory

#caminho = askdirectory()
#print(caminho)
path = '/home/oziel/Documentos/Alunos/PauloB/data/today'

lista_caminho = os.listdir(path)
print(lista_caminho)
erro_geral = 0
contador_geral = 0
tool = lt.LanguageTool("pt-BR")
for caminho in lista_caminho:
    print("*-" * 24)
    print("Fase 1")
    print("Total: ", len(lista_caminho))
    print("Do total estamos na pasta: ", contador_geral + 1)
    print("Analisando: ", caminho)
    path1 = path + "/" + caminho
    caminho_pdf = os.listdir(path1)
    #print(caminho_pdf)
    if len(caminho_pdf) != 0:
        erro_local = 0
        contador = 0
        for caminho1 in caminho_pdf:
            print("*-" * 24)
            print("Fase 2 - Analisando ", len(caminho_pdf))
            print("Arquivo atual: ", caminho1)
            print("Contador: ", contador + 1)
            print((round((contador+1)/len(caminho_pdf),2))*100,"%")
            
            titulo = caminho1.replace(".pdf",'') + '.txt'
            path_relatorio = path1 + "/Erro_" + titulo
            
            print(path_relatorio)
            
            #path2 = path1 + "/" + caminho1
            #output_string = StringIO()
            #with open(path2,'rb') as in_file:
            #    parser = PDFParser(in_file)
            #    doc = PDFDocument(parser)
            #    rsrcmgr = PDFResourceManager()
            #    device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
            #    interpreter = PDFPageInterpreter(rsrcmgr,device)
            #    for page in PDFPage.create_pages(doc):
            #        interpreter.process_page(page)
            #    print("Extração de texto completa!")
            #    text = output_string.getvalue
            contador += 1
        contador_geral += 1
    else:
        print(caminho, "não contém arquivo(s)")
        print("Total: ", len(lista_caminho))
        print("Do total estamos na pasta: ", contador_geral + 1)
        contador_geral += 1