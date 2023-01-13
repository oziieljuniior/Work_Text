'''
created: Oziel Ramos
contat: ozieljr14@gmail.com
'''


#Biblioteca suporte para extração de texto, reconhecimento do texto como string
from io import StringIO

#Biblioteca para extração de texto
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

#Biblioteca que realiza análise textual
import language_tool_python as lt

#Bibliotecas que configuram os caminhos pelo pc
import os
#from tkinter.filedialog import askdirectory

#Blioteca que gerencia planilha
from openpyxl import Workbook

#2 planilhas foram criadas: Geral e Local
excel1 = Workbook()
data_geral =  excel1.active
data_geral.title = "Data_Geral"
data_geral['A1'] = 'Qt_Total'
excel2 = Workbook()
data_local = excel2.active
data_local.title = "Empresas"
data_local["A1"] = "Empresas"
data_local["B1"] = "Qt_Erros"
indice_excel = 1

#Configurar caminho
#caminho = askdirectory()
#print(caminho)
#'/home/oziel/Documentos/Alunos/PauloB/data/today'
path = 'C:\\Users\\Riallen\\Documents\\Att\\treinamento1'
lista_caminho = os.listdir(path)
print(lista_caminho)

#Contadores iniciais para análise de erro
erro_geral = 0
contador_geral = 0

#Configuração da biblioteca para análise textual 
tool = lt.LanguageTool("pt-BR")

#Percorre e lista arquivos na pasta geral
#caminho é a empresa
for caminho in lista_caminho:
    #configuração para acompanhamento no terminal
    print("*-" * 24)
    print("Fase 1")
    print("Total: ", len(lista_caminho))
    print("Do total estamos na pasta: ", contador_geral + 1)
    print("Analisando: ", caminho)
    
    #Configuração de caminho para percorrer pastas das empresas
    path1 = path + "\\" + caminho
    caminho_pdf = os.listdir(path1)
    #print(caminho_pdf)

    #Contador dos erros locais de cada empresa
    erro_local = 0
    #Condicional para realizar a captura e análise textual
    if len(caminho_pdf) != 0:
        #contador qualquer
        contador = 0
        for caminho1 in caminho_pdf:
            #configuração para acomponhar análise pelo terminal
            print("*-" * 24)
            print("Fase 2 - Analisando ", len(caminho_pdf))
            print("Arquivo atual: ", caminho1)
            print("Contador: ", contador + 1)
            print((round((contador+1)/len(caminho_pdf),2))*100,"%")
            
            #Ajuste para criação de relatório local de cada empresa e ano
            titulo = caminho1.replace(".pdf",".txt")
            path_relatorio = path1 + "\\Erro_" + titulo
            print(path_relatorio)
            
            #Configuração para extração do texto
            path2 = path1 + "\\" + caminho1
            output_string = StringIO()
            with open(path2,'rb') as in_file:
                parser = PDFParser(in_file)
                doc = PDFDocument(parser)
                rsrcmgr = PDFResourceManager()
                device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
                interpreter = PDFPageInterpreter(rsrcmgr,device)
                for page in PDFPage.create_pages(doc):
                    interpreter.process_page(page)
            #variavel de retorno da extração do texto
            text = output_string.getvalue()
            print("Extração de texto completa!")
            
            #Análise textual
            matches = tool.check(text)
            erros_apontados = len(matches)
            #criação de arquivo txt para relatório local 
            aperro = open(path_relatorio, 'x', encoding='utf-8')
            for i in range(0, erros_apontados):
                if matches[i].ruleId != 'WHITESPACE_RULE' and matches[i].ruleId != 'DASH_RULE':
                    erro_local = erro_local + 1
                    aperro.write(str(erro_local))
                    aperro.write(str(matches[i]))
            aperro.close()
            print("Análise realizada com sucesso!")
            print("Relatório criado com sucesso")

            erro_local = erro_local + erros_apontados
            contador += 1
        indice_planilhaA = 'A' + str(indice_excel)
        indice_planilhaB = 'B' + str(indice_excel)
        data_local[indice_planilhaA] = caminho
        data_local[indice_planilhaB] = erro_local
        print("Planilha local atualizada")
        
        erro_geral = erro_geral + erro_local
        indice_excel += 1
        contador_geral += 1
    else:
        print("*-" * 24)
        print("Fase 2 - Analisando ", len(caminho_pdf))
        print(caminho, "não contém arquivo(s)")
        print("Do total estamos na pasta: ", contador_geral + 1)
        
        indice_planilhaA = 'A' + str(indice_excel)
        indice_planilhaB = 'B' + str(indice_excel)
        data_local[indice_planilhaA] = caminho
        data_local[indice_planilhaB] = erro_local
        
        erro_geral = erro_geral + erro_local
        indice_excel += 1
        contador_geral += 1
excel1.save("Data_Geral.xlsx")
excel2.save("Data_Local.xlsx")