import os
import pyautogui
import time
import shutil

lista = ['TEGA3', 'TTEN3', 'VVEO3', 'WEST3']
#print(lista)
path = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data'
path_final = os.listdir(path)
#print(path_final)
for name in path_final:
    for name1 in lista:
        if name == name1:
            path_empresas = path + '\\' + name
            relatorios_listados = os.listdir(path_empresas)
            #print(relatorios_listados)
            anos = []
            for ano in relatorios_listados:
                c = ano.split("_")
                #print(c)
                c0 = c[1].split(".")
                #print(c0)
                anos.append(int(c0[0]))
            print(anos)
            i = 0
            while i == 0:
                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\firefox.png') != None:
                    print("Fase 1")
                    pyautogui.hotkey('ctrl' , 't')
                    time.sleep(2)
                    pyautogui.write('https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm', interval = 0.05)
                    pyautogui.press("enter")
                    time.sleep(10)
                    j = 0
                    while j == 0:
                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\digitar_nome.png') != None:
                            print("Fase 2")
                            k = 0
                            time.sleep(10)
                            pyautogui.doubleClick('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\digitar_nome.png', interval = 2)
                            pyautogui.write(name, interval = 0.5)
                            pyautogui.press('enter')
                            time.sleep(5)
                            pyautogui.click((238,631), interval = 1)
                            while k == 0:
                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\sobre_empresa.png') != None:
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\sobre_empresa.png', interval = 2)
                                    k += 1
                            k = 0
                            while k == 0:
                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\relatorio_estruturado.png') != None:
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\relatorio_estruturado.png', interval = 2)
                                    k += 1
                            k = 0
                            atua = [2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021]
                            print("Fase 3")
                            print("procura")
                            for name2 in anos:
                                for name1 in atua:
                                    if name2 == name1:
                                        atua.remove(name2)
                            atua.append(2022)
                            print(atua)
                            print("Fase 3")
                            for name1 in atua:
                                time.sleep(5)
                                print(name1)
                                k = 0
                                print(k)
                                if name1 == 2012:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2012.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        
                                                    k += 1
                                                    l += 1

                                if name1 == 2013:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2013.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1

                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2014:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2014.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2015:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2015.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2016:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2016.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2017:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2017.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2018:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2018.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2019:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2019.png', interval = 2)
                                    k = 0
                                    print(k)
                                
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2020:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2020.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1

                                if name1 == 2021:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2021.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                        

                                                    k += 1
                                                    l += 1                                                                                

                                if name1 == 2022:
                                    pyautogui.click((143,537), interval = 2)
                                    time.sleep(3)
                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\2022.png', interval = 2)
                                    k = 0
                                    print(k)
                                    
                                    while k == 0:
                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png') != None:
                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\info_tri.png', interval = 1)
                                            time.sleep(15)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Consolidades.png', interval = 2)
                                                    l = l + 1
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\DFs_Individuais.png', interval = 2)
                                                    l = l + 1
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\notas_explicativas.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                            l = 0
                                            while l == 0:
                                                if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png') != None:
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\download.png', interval = 2)
                                                    l = l + 1
                                                    time.sleep(5)
                                                    name_pdf = name + "_" + str(name1)
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.click('C:\\Users\\Riallen\\Documents\\Work_Text\\Point_Click\\salvar.png', interval = 1)
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')
                                                    time.sleep(5)
                                                    pyautogui.hotkey('ctrl', 'w')

                                                
                                        
                    
                            
                                                    i += 1
                                                    j += 1
                                                    k += 1
                                                    l += 1
                                
                                                        
                            path_download =   "C:\\Users\\Riallen\\Downloads"
                            lista_download = os.listdir(path_download)
                            path1 = 'C:\\Users\\Riallen\\Documents\\Work_Text\\Data\\' + name
                            print(path1)
                            for name3 in lista_download:
                                source = path_download+ '\\' + name3
                                print(source)
                                shutil.move(source, path1)

            