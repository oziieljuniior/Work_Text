import os

caminho = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data_Final'
path_data_final = os.listdir(caminho)
print(path_data_final)

for name in path_data_final:
    caminho1 = caminho + "\\" + name
    path_empresa = os.listdir(caminho1)
    #print(path_empresa)
    for name1 in path_empresa:
        texto = name1.split("_")
        texto1 = texto[1].split(".")
        #print(int(texto1[0]) - 1)
        ano = int(texto1[0]) - 1
        name_final = texto[0] + "_" + str(ano) + ".pdf"
        #print(name_final)
        path_namefinal = caminho1 + "\\" + name_final
        #print(path_namefinal)
        path_nameinicial = caminho1 + "\\" + name1
        os.rename(path_nameinicial, path_namefinal)