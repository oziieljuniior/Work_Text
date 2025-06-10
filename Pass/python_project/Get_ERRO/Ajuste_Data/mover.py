import os

caminho = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data_Final'
caminho_ano = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data_Final_Ano'

#path_data_final = os.listdir(caminho)
#print(path_data_final)
#ida
#for name in path_data_final:
#    caminho1 = caminho + "\\" + name
#    path_empresa = os.listdir(caminho1)
#    #print(path_empresa)
#    for name1 in path_empresa:
#        path_geral = caminho1 + "\\" + name1
#        caminho_final = caminho_ano + "\\" + name1
#        os.rename(path_geral, caminho_final)

path_dataano = os.listdir(caminho_ano)
#print(path_dataano)

ano = ['2010','2011','2012','2013','2014','2015','2016','2017','2018','2019','2020','2021']

for name in ano:
    pasta_ano = caminho + "\\" + name
    #os.makedirs(pasta_ano)
    for name1 in path_dataano:
        texto = name1.split("_")
        texto1 = texto[1].split(".")
        #print(texto1)
        if texto1[0] == name:
            pasta_final = pasta_ano + "\\" + name1
            pasta_inicial = caminho_ano + "\\" + name1
            #print(pasta_final)
            #print(pasta_inicial)
            os.rename(pasta_inicial, pasta_final)