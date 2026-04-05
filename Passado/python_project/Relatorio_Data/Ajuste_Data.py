import os

#path data
#C:\Users\Riallen\Documents\Work_Text\Data
path = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data'
path_final = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data_Final'
list_dir = os.listdir(path)
i = 0
for name in list_dir:
    #print(name)
    #print(len(name))
    path_pasta = path + '\\' + name
    list_rel = os.listdir(path_pasta)
    #print(list_rel)
    if len(list_rel) == 12 or len(list_rel) == 11:
        #print(list_rel)
        #print(len(list_rel))
        path_pasta_nova = path_final + '\\' + name
        os.mkdir(path_pasta_nova)
        for name1 in list_rel:
            print(name1)
            path_saida = path_pasta + '\\' + name1
            path_chegada = path_pasta_nova + '\\' + name1
            os.rename(path_saida, path_chegada)
        i += 1

print(i)