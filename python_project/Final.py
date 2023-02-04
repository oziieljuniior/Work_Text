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
from tkinter.filedialog import askdirectory

#Blioteca que gerencia planilha
from openpyxl import Workbook
import pandas as pd 


import datetime as dt
inicio = dt.datetime.now().strftime("%H:%M:%S")

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


#Configuração da biblioteca para análise textual 
tool = lt.LanguageTool("pt-BR")
        


#Configurar caminho
path = askdirectory(title = 'Caminho Data salva')
path_resultado = askdirectory(title = 'Caminho onde o resultado da pesquisa deve ser salvo')
path_saida_original = askdirectory(title = 'Caminho de Saida')
path_texto = askdirectory(title = 'Caminho Texto')
#path.replace("/","\\")
#print(path)
#'/home/oziel/Documentos/Alunos/PauloB/data/today'
#path = 'C:\\Users\\Riallen\\Documents\\Att\\treinamento5 '
lista_caminho = os.listdir(path)
print(lista_caminho)

#Contadores iniciais para análise de erro
erro_geral = 0
contador_geral = 0

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
    path1 = path + "/" + caminho
    caminho_pdf = os.listdir(path1)
    #print(caminho_pdf)

    #Contador dos erros locais de cada empresa
    erro_local = 0
    #Condicional para realizar a captura e análise textual
    if len(caminho_pdf) != 0:
        #contador qualquer
        contador = 0
        excel3 = Workbook()
        data_local_ano = excel3.active
        data_local_ano.title = "Empresas_Anos"
        indice_excel1 = 1

        for caminho1 in caminho_pdf:
            #configuração para acomponhar análise pelo terminal
            print("*-" * 24)
            print("Fase 2 - Analisando ", len(caminho_pdf))
            print("Arquivo atual: ", caminho1)
            print("Contador: ", contador + 1)
            print((round((contador+1)/len(caminho_pdf),2))*100,"%")
            
            #Ajuste para criação de relatório local de cada empresa e ano
            titulo = caminho1.replace(".pdf",".txt")
            path_relatorio = path_resultado + "/Erro_" + titulo
            print(path_relatorio)
            
            #configurando texto txt
            path_texto_original = path_texto +"/"+ titulo
            texto_txt = open(path_texto_original, 'x', encoding = 'utf-8')
            
            #Configuração para extração do texto
            path2 = path1 + "/" + caminho1
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
            texto_txt.write(text)
            print("Extração de texto completa!")
            texto_txt.close()
            
            

            #Análise textual
            print("Analisando erros")
            print(dt.datetime.now().strftime("%H:%M:%S"))
            matches = tool.check(text)
            erros_apontados = len(matches)
            #criação de arquivo txt para relatório local 
            aperro = open(path_relatorio, 'x', encoding='utf-8')
            
            erro_anual = 0
            
            for i in range(0, erros_apontados):
                #PORTUGUESE_WORD_REPEAT_BEGINNING_RULE -> Verificar possibilidade de retirar números do relatório
                if matches[i].ruleId != 'WHITESPACE_RULE' and matches[i].ruleId != 'DASH_RULE' and matches[i].ruleId != 'HUNSPELL_RULE' and matches[i].ruleId != 'SPACE_BEFORE_PUNCTUATION' and matches[i].ruleId != 'ORDINAL_ABREVIATION' and matches[i].ruleId != 'SENT_START_NUM' and matches[i].ruleId != 'ROMAN_NUMBERS_CHECKER' and matches[i].ruleId != 'DECIMAL_COMMA' and matches[i].ruleId != 'PT_COMPOUNDS_POST_REFORM' and matches[i].ruleId != 'UNPAIRED_BRACKETS' and matches[i].ruleId != 'PT_BARBARISMS_REPLACE' and matches[i].ruleId != 'PT_BR_SIMPLE_REPLACE' and matches[i].ruleId != 'CHEMICAL_FORMULAS_TYPOGRAPHY' and matches[i].ruleId != 'GENERAL_NUMBER_FORMAT' and matches[i].ruleId != 'GENERAL_NUMBER_AGREEMENT_ERRORS'and matches[i].ruleId != 'PHRASE_REPETITION' and  matches[i].ruleId !='ABREVIATIONS_PUNCTUATION' and matches[i].ruleId != 'COPYRIGHT' and matches[i].ruleId != 'BARBARISMS' and matches[i].ruleId != 'UPPERCASE_AFTER_COMMA' and matches[i].ruleId != 'GENERAL_GENDER_AGREEMENT_ERRORS' and matches[i].ruleId != 'ZERO_IN_DAYS_OF_THE_MONTH' and matches[i].ruleId != 'PORTUGUESE_WORD_REPEAT_RULE' and matches[i].ruleId != 'GENTILICOS_LINGUAS' and matches[i].ruleId != 'LOWERCASE_RARE_WORDS_GEOGRAPHICAL' and matches[i].ruleId != 'UPPERCASE_SENTENCE_START' and matches[i].ruleId != 'REPEATED_WORDS_3X' and matches[i].ruleId != 'SI_UNITS_CASE' and matches[i].ruleId != 'PERCENT_WITHOUT_SPACE' and matches[i].ruleId != 'ARTICLES_PRECEDING_LOCATIONS' and matches[i].ruleId != 'UNPAIRED_BRACKET_SUGGESTIONS' and matches[i].ruleId != 'COMMA_PARENTHESIS_WHITESPACE' and matches[i].ruleId != 'DOUBLE_PUNCTUATION' and matches[i].ruleId != 'A_WORD' and matches[i].ruleId != 'HOMONYM_ASSO_10' and matches[i].ruleId != 'AO90_MONTHS_CASING' and matches[i].ruleId != 'ABREVIATIONS_PUNCTUATION' and matches[i].ruleId != 'NO_SPACE_CLOSING_QUOTE' and matches[i].ruleId != 'CERCA_DE_NR' and matches[i].ruleId != 'NUMBER_ABREVIATION' and matches[i].ruleId != 'INFORMALITIES' and matches[i].ruleId != 'ENUMERATIONS_AND_AND' and matches[i].ruleId != 'ARCHAISMS' and matches[i].ruleId != 'INVALID_DATE' and matches[i].ruleId != 'PT_CLICHE_REPLACE' and matches[i].ruleId != 'QUE_SER_ESTAR_PARTPASSADO' and matches[i].ruleId != 'ALTERNATIVE_CONJUNCTIONS_COMMA' and matches[i].ruleId != 'E_É_SÃO_FOI_FORAM_SENDO_SIDO' and matches[i].ruleId != 'QUE_FORAM_FOI_SÃO_É_SENDO' and matches[i].ruleId != 'INTERNET_ABBREVIATIONS' and matches[i].ruleId != 'PORTUGUESE_WORD_REPEAT_BEGINNING' and matches[i].ruleId != 'INTERJECTIONS_PUNTUATION' and matches[i].ruleId !='ERRO_DE_CONCORDNCIA_DO_GÉNERO_MASCULINO_O' and matches[i].ruleId != 'FRAGMENT_TWO_PREPOSTIONS' and matches[i].ruleId != 'ADVERBIOS_MODO_EM_SEQUENCIA' and matches[i].ruleId != 'CUJO_LIGACAO_NOME_ADJETIVO_NUMERAL' and matches[i].ruleId != 'PP_OBJ_IND' and matches[i].ruleId != 'PT_SIMPLE_REPLACE' and matches[i].ruleId != 'REDUNDANT_CONJUNCTIONS' and matches[i].ruleId != 'COLOCAÇÃO_ADVÉRBIO' and matches[i].ruleId != 'INTERROGATIVES_PUNTUATION' and matches[i].ruleId != 'YEAR_NUMBER_FORMAT' and matches[i].ruleId != 'VERB_COMMA_CONJUNCTION' and matches[i].ruleId != 'REDUNDANCY_JUNTO_COM' and matches[i].ruleId != 'PORTUGUESE_WRONG_WORD_IN_CONTEXT' and matches[i].ruleId != 'PT_WEASELWORD_REPLACE' and matches[i].ruleId != 'COLOCACAO_ADVERBIOS_LUGAR' and matches[i].ruleId != 'SENTENCE_WHITESPACE' and matches[i].ruleId != 'DATE_NEW_YEAR' and matches[i].ruleId != 'kWh' and matches[i].ruleId != 'ALÉM_AQUÉM_RECÉM' and matches[i].ruleId != 'PROFANITY' and matches[i].ruleId != 'PARENTESESE_AND_QUOTES_SPACING' and matches[i].ruleId != 'EM_LONGO_PRAZO' and matches[i].ruleId != 'SUBSTANTIVO_PLURAL_E_CHAVE':
                    erro_local = erro_local + 1
                    erro_anual = erro_anual + 1
                    aperro.write("Erro: ")
                    aperro.write(str(erro_local))
                    aperro.write('\n')
                    aperro.write(str(matches[i]))
                    aperro.write('\n')
            aperro.close()
            
            indice_planilhaeA = "A" + str(indice_excel1)
            indice_planilhaeB = "B" + str(indice_excel1)
            data_local_ano[indice_planilhaeA] = caminho1
            data_local_ano[indice_planilhaeB] = erro_anual
            indice_excel1 += 1
            
            print("Análise realizada com sucesso!")
            print("Relatório criado com sucesso")
            
            contador += 1
        
        name = path_resultado + "/" + caminho + ".xlsx"
        print("Salvo: ", name)
        excel3.save(name)
        
        indice_excel += 1
        indice_planilhaA = 'A' + str(indice_excel)
        indice_planilhaB = 'B' + str(indice_excel)
        data_local[indice_planilhaA] = caminho
        data_local[indice_planilhaB] = erro_local
        print("Planilha local atualizada")
        print("Quantidade de erros localizadas: ", erro_local)

       

        erro_geral = erro_geral + erro_local
        contador_geral += 1

    
    else:
        print("*-" * 24)
        print("Fase 2 - Analisando ", len(caminho_pdf))
        print(caminho, "não contém arquivo(s)")
        print("Do total estamos na pasta: ", contador_geral + 1)
        
        indice_excel += 1
        indice_planilhaA = 'A' + str(indice_excel)
        indice_planilhaB = 'B' + str(indice_excel)
        data_local[indice_planilhaA] = caminho
        data_local[indice_planilhaB] = erro_local
        
        erro_geral = erro_geral + erro_local
        contador_geral += 1
    
    print("Realizando ajuste de pasta: ", path1)
    path_saida = path_saida_original + "/" + caminho
    print("Para: ", path_saida)
    os.rename(path1, path_saida)


excel2.save(path_resultado + "/Data_Local.xlsx")
print("Calculando quantidade total de erros")
a = []
data = pd.read_excel(path_resultado + "/Data_Local.xlsx")
#print(len(data))
#print(data.columns)
for i in range(0,len(data)):
    print(data['Qt_Erros'][i])
    a.append(data['Qt_Erros'][i])
qt_total = sum(a)
data_geral['B1'] = qt_total

excel1.save(path_resultado + "/Data_Geral.xlsx")

fim = dt.datetime.now().strftime("%H:%M:%S")


print("Consulta finalizada com sucesso")
print("Processo começou: ", inicio)
print("Processo finalizou: ", fim)