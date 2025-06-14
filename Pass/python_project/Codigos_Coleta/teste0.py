import pyautogui
import time
from PIL import Image
import pandas as pd
import os #Busca e acesso das pastas do S.O.
import shutil

data = pd.read_excel('C:/Users/Riallen/Documents/Att/Cias não financeiras - 12-12-2022.xlsx')
consulta = data['ticker'].to_list()
t1 = len(consulta)
for emp in range(0,t1):
    path = "C:/Users/Riallen/Documents/Att/Data/" + consulta[emp]
    os.makedirs(path)
    
    i = 0
    while i == 0:
        print("Fase 1")
        im1 = pyautogui.screenshot().save("im1.png")
        im2 = Image.open("im1.png")
        #[(1, (249, 249, 251))] [(1, (249, 249, 251))]
        #[(1, (251, 251, 251))] [(1, (249, 249, 251))]
        area1 = (169,68,170,69)
        area2 = (188,208,189,209)
        im3_corte1 = im2.crop(area1)
        im4_corte2 = im2.crop(area2)
        im3_corte1.save('foto01.jpeg','jpeg')
        im4_corte2.save('foto02.jpeg','jpeg')
        im5 = Image.open('foto01.jpeg')
        im6 = Image.open('foto02.jpeg')
        convert_im1 = im5.convert('RGB').getcolors()
        convert_im2 = im6.convert('RGB').getcolors()
        print(convert_im1,convert_im2)
        
        if convert_im1 == [(1, (249, 249, 251))] and convert_im2 == [(1, (249, 249, 251))]:
            print("Fase 2")
            j = 0
            pyautogui.hotkey('ctrl','t')
            time.sleep(3)
            #https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm
            pyautogui.write('https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm', interval = 0.10)
            pyautogui.press('enter')
            time.sleep(10)
            while j == 0:
                print("Frase 3")
                im7 = pyautogui.screenshot().save("im2.png")
                im8 = Image.open("im2.png")
                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                #[(1, (2, 42, 104))] [(1, (255, 255, 255))]
                #[(1, (1, 44, 99))] [(1, (223, 223, 223))]
                #[(1, (255, 255, 255))] [(1, (223, 223, 223))]
                area3 = (136,289,137,290)
                area4 = (323,679,324,680)
                im9_corte1 = im8.crop(area3)
                im10_corte2 = im8.crop(area4)
                im9_corte1.save('foto03.jpeg','jpeg')
                im10_corte2.save('foto04.jpeg','jpeg')
                im11 = Image.open('foto03.jpeg')
                im12 = Image.open('foto04.jpeg')
                convert_im3 = im11.convert('RGB').getcolors()
                convert_im4 = im12.convert('RGB').getcolors()
                print(convert_im3,convert_im4)
                if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (223, 223, 223))]:
                    print("Fase 4")
                    k = 0
                    time.sleep(10)
                    pyautogui.doubleClick(272,621)
                    pyautogui.write(consulta[emp], interval = 0.5)
                    pyautogui.press('enter')
# =============================================================================
#                     time.sleep(5)
# =============================================================================
                    while k == 0:
                        print("Fase 5")
                        
                        im13 = pyautogui.screenshot().save("im3.jpeg")
                        im14 = Image.open("im3.jpeg")
                        #[(1, (255, 255, 255))]
                        #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                        #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                        #[(1, (1, 44, 99))] [(1, (223, 223, 223))]
                        #[(1, (255, 255, 255))] [(1, (223, 223, 223))]
                        area5 = (272,307,273,308)
                        area6 = (493,687,494,688)
                        im15_corte1 = im14.crop(area5)
                        im16_corte2 = im14.crop(area6)
                        im15_corte1.save('foto05.jpeg','jpeg')
                        im16_corte2.save('foto06.jpeg','jpeg')
                        im17 = Image.open('foto05.jpeg')
                        im18 = Image.open('foto06.jpeg')
                        convert_im5 = im11.convert('RGB').getcolors()
                        convert_im6 = im12.convert('RGB').getcolors()
                        print(convert_im5,convert_im6)
                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (223, 223, 223))]:
                            time.sleep(3)
                            l = 0
                            print("Fase 6")
                            time.sleep(10)
                            pyautogui.doubleClick(207,650)
                            #possivel colocar outra captura
                            time.sleep(5)
                            pyautogui.click(1133,467)
                            time.sleep(3)
                            pyautogui.click(1067,527)
                            while l == 0:
                                print("Fase 7")
                                im13 =  pyautogui.screenshot().save("im3.jpeg")
                                im14 = Image.open("im3.jpeg")
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                #[(1, (1, 44, 99))] [(1, (255, 255, 255))]
                                #[(1, (1, 44, 99))] [(1, (223, 223, 223))]
                                area5 = (725,310 ,726,311)
                                area6 = (175,570,176,571)
                                im15_corte1 = im14.crop(area5)
                                im16_corte2 = im14.crop(area6)
                                im15_corte1.save('foto05.jpeg','jpeg')
                                im16_corte2.save('foto06.jpeg','jpeg')
                                im17 = Image.open('foto05.jpeg')
                                im18 = Image.open('foto06.jpeg')
                                convert_im5 = im11.convert('RGB').getcolors()
                                convert_im6 = im12.convert('RGB').getcolors()
                                print(convert_im5,convert_im6)
                                if convert_im3 == [(1, (1, 44, 99))] and convert_im4 == [(1, (223, 223, 223))]:
                                    print("Fase 8")
                                    time.sleep(10)
                                    pyautogui.click(155,540)
                                    #
                                    
                                    
                                    m = 0
                                    while m <= 11:
                                        cot = 2010 + m
                                        ano = str(cot)
                                        if cot == 2010:
                                            n = 0
                                            print("Ano ", ano)
                                            time.sleep(3)
# =============================================================================
#                                             pyautogui.press('pageup')
# =============================================================================
                                            pyautogui.click(133,425)                                            
                                            time.sleep(3)
                                            #pyautogui.click((1897,89),interval = 1)
                                            #time.sleep(1)
                                            #pyautogui.click((1699,524),interval = 1)
                                            #pyautogui.press('enter')
                                            
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao ')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            #/html/body/app-root/app-companies-menu-select/div/app-companies-structured-reports/form/div[2]/div/div[2]/div/div[1]/div[1]/div[2]/p/a
                                            #/html/body/app-root/app-companies-menu-select/div/app-companies-structured-reports/form/div[2]/div/div[2]/div/div[1]/div[1]/div[2]/p/a
                                            
                                            m += 1
                                        
                                        elif cot == 2011:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,412)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        
                                        
                                        elif cot == 2012:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,381)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2013:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,359)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2014:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,335)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2015:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,310)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2016:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,284)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            m += 1
                                        elif cot == 2017:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,254)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                    
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2018:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,229)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2019:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,205)
                                            
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2020:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,181)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2021:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((1257,705), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(155,540)
                                            time.sleep(3)
                                            pyautogui.click(133,152)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                x = pyautogui.locateOnScreen('C:/Users/Riallen/Documents/Att/point.png')
                                                print(x)
                                                if x != None:
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    pyautogui.click('C:/Users/Riallen/Documents/Att/point.png')
                                                
                                                    time.sleep(15)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\DFs_Consolidades.png', interval = 2)
                                                            y = y + 1
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\notas_explicativas.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    y = 0
                                                    while y == 0:
                                                        if pyautogui.locateOnScreen('C:\\Users\\Riallen\\Documents\\Att\\download.png') != None:
                                                            pyautogui.click('C:\\Users\\Riallen\\Documents\\Att\\download.png', interval = 2)
                                                            y = y + 1
                                                    time.sleep(5)
                                                    name_pdf = consulta[emp] + "_" + cot
                                                    print(name_pdf)
                                                    pyautogui.write(name_pdf, interval = 0.5)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)
                                                    pyautogui.click((701,22), interval = 2)
                                                    pyautogui.click(3,365)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.click(3,365)
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            
                                            m += 1
                                            i += 1
                                            j += 1
                                            k += 1
                                            l += 1
                                    
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                            
                                        
                                    
                                        
                                    #pyautogui.hotkey('ctrl','w')
                                        
                                print("Conversão realizada")
                                pyautogui.hotkey('ctrl','w')
                                p = 0
                                while p == 0:
                                    #C:/Users/M&A SOLUTIONS BRASIL/Downloads
                                    path1 = os.listdir("C:/Users/Riallen/Downloads")
                                    for arquivo in path1:
                                        source = "C:/Users/Riallen/Downloads/" + arquivo
                                        shutil.move(source, path)
                                    p = p +1
        