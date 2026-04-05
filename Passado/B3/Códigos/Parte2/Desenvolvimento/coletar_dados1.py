import pyautogui
import pyperclip
import time
import pandas as pd
from PIL import Image
import pyscreenshot as ImageGrab
import os
import shutil

#E:\Python_Project\CNFB3.xlsx
data = pd.read_excel('E:\Python_Project\CNFB3.xlsx')
consulta = data['TICKER'].to_list()
t1 = len(consulta)
#RRP3 ~ não consta]
text = 'Formulário de Referência'
text1 = 'Comentários dos diretores'
text2 = 'Gerenciamento de riscos e controles internos'
text3 = 'Todos'

#Contar o valor inicial de (2,t1). Isso é determinado pela quantidade de arquivos que está na pasta Data\Empresas
for emp in range(38,t1):
    #E:\Python_Project\Data\Empresas
    path = 'E:\Python_Project\Data\Empresas\ ' + consulta[emp]
    print(consulta[emp])
    os.makedirs(path)
    
    i = 0
    while i == 0:
        print("Fase 1")
        im1 = ImageGrab.grab().save("im1.png")
        im2 = Image.open("im1.png")
        
        area1 = (1027, 21, 1028, 22)
        area2 = (1000, 64, 1001, 65)
        im3_corte1 = im2.crop(area1)
        im4_corte2 = im2.crop(area2)
        im3_corte1.save('foto01.jpeg','jpeg')
        im4_corte2.save('foto02.jpeg','jpeg')
        im5 = Image.open('foto01.jpeg')
        im6 = Image.open('foto02.jpeg')
        convert_im1 = im5.convert('RGB').getcolors()
        convert_im2 = im6.convert('RGB').getcolors()
        print(convert_im1,convert_im2)
        
        if convert_im1 == [(1, (211, 227, 253))] and convert_im2 == [(1, (235, 242, 250))]:
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
                im7 = ImageGrab.grab().save("im2.png")
                im8 = Image.open("im2.png")
                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                area3 = (575, 775, 576, 776)
                area4 = (1245, 775, 1246, 776)
                im9_corte1 = im8.crop(area3)
                im10_corte2 = im8.crop(area4)
                im9_corte1.save('foto03.jpeg','jpeg')
                im10_corte2.save('foto04.jpeg','jpeg')
                im11 = Image.open('foto03.jpeg')
                im12 = Image.open('foto04.jpeg')
                convert_im3 = im11.convert('RGB').getcolors()
                convert_im4 = im12.convert('RGB').getcolors()
                print(convert_im3,convert_im4)
                #[(1, (88, 204, 241))] [(1, (255, 255, 255))]
                #[(1, (255, 255, 255))] [(1, (88, 204, 241))]
                if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (88, 204, 241))]:
                    print("Fase 4")
                    k = 0
                    time.sleep(10)
                    pyautogui.doubleClick(455,775)
                    pyautogui.write(consulta[emp], interval = 0.5)
                    pyautogui.press('enter')
# =============================================================================
#                     time.sleep(5)
# =============================================================================
                    while k == 0:
                        print("Fase 5")
                        
                        im13 = ImageGrab.grab().save("im3.jpeg")
                        im14 = Image.open("im3.jpeg")
                        #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                        area5 = (770, 360, 771, 361)
                        area6 = (455, 775, 456, 776)
                        im15_corte1 = im14.crop(area5)
                        im16_corte2 = im14.crop(area6)
                        im15_corte1.save('foto05.jpeg','jpeg')
                        im16_corte2.save('foto06.jpeg','jpeg')
                        im17 = Image.open('foto05.jpeg')
                        im18 = Image.open('foto06.jpeg')
                        convert_im5 = im11.convert('RGB').getcolors()
                        convert_im6 = im12.convert('RGB').getcolors()
                        print(convert_im5,convert_im6)
                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (88, 204, 241))]:
                            time.sleep(3)
                            l = 0
                            print("Fase 6")
                            time.sleep(10)
                            pyautogui.doubleClick(455,775)
                            #possivel colocar outra captura
                            time.sleep(3)
                            pyautogui.click(1630,575)
                            time.sleep(3)
                            pyautogui.click(1630,650)
                            while l == 0:
                                print("Fase 7")
                                im13 = ImageGrab.grab().save("im3.jpeg")
                                im14 = Image.open("im3.jpeg")
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                area5 = (1200,310,1201,311)
                                area6 = (300,680,301,681)
                                im15_corte1 = im14.crop(area5)
                                im16_corte2 = im14.crop(area6)
                                im15_corte1.save('foto05.jpeg','jpeg')
                                im16_corte2.save('foto06.jpeg','jpeg')
                                im17 = Image.open('foto05.jpeg')                                                                                                                                                                                                                            
                                im18 = Image.open('foto06.jpeg')
                                convert_im5 = im11.convert('RGB').getcolors()
                                convert_im6 = im12.convert('RGB').getcolors()
                                print(convert_im5,convert_im6)
                                if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (88, 204, 241))]:
                                    print("Fase 8")
                                    time.sleep(10)
                                    pyautogui.click((300,680), interval=1)
                                    time.sleep(1)
                                    #
                                               
                                    m = 0
                                    while m <= 14:
                                        cot = 2010 + m
                                        ano = str(cot)
                                        if cot == 2010:
                                            n = 0
                                            print("Ano ", ano)
                                            time.sleep(3)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            time.sleep(2)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                #[(1, (236, 76, 44))] [(1, (232, 85, 33))]
                                                #[(1, (71, 209, 124))] [(1, (54, 213, 121))]
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                     
                                                    pyautogui.click((405, 574), interval = 1)
                                                    
                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        #(409, 133)
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')

                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)
                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1 
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
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
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)
                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        
                                        
                                        elif cot == 2012:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2013:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2014:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2015:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                                    
                                            m += 1
                                        elif cot == 2016:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            m += 1
                                        elif cot == 2017:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2018:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2019:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2020:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    #(596, 394)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776,641,777,642)

                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
    
                                            
                                            m += 1
                                        elif cot == 2021:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
                                            m += 1
                                        elif cot == 2022:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
                                            m += 1
                                        elif cot == 2023:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    

                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    n += 1
                                                    time.sleep(3)
                                            m += 1
                                        elif cot == 2024:
                                            n = 0
                                            print("Ano ", ano)
                                            pyautogui.click((300,680), interval=1)
                                            time.sleep(1)
                                            pyautogui.write(ano, interval = 0.5)
                                            pyautogui.press('enter')    
                                            
                                            time.sleep(10)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyperclip.copy(text)
                                            pyautogui.hotkey('ctrl', 'v')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (358, 512, 359, 513)
                                                area4 = (405, 574, 406, 575)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('foto03.jpeg')
                                                im12 = Image.open('foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if (convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]) or (convert_im3 == [(1, (254, 151, 46))] and convert_im4 == [(1, (250, 255, 3))]):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((405, 574), interval = 1)

                                                    time.sleep(10)
                                                    a1 = 0
                                                    while a1 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (901, 133, 902, 134)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (14, 45, 89))] or convert_im3 == [(1, (234, 255, 255))] or convert_im3 == [(1, (0, 70, 132))]:
                                                            print("click on the bellow")
                                                            pyautogui.click((1785, 184), interval = 1)
                                                            a1 = 1
                                                
                                                    time.sleep(15)
                                                    pyautogui.click((596, 393), interval = 1)
                                                    time.sleep(2)
                                                    
                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text1)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    
                                                    time.sleep(2)
                                                    #[(1, (255, 157, 40))]
                                                    im19 = ImageGrab.grab().save("im19.jpeg")
                                                    im20 = Image.open("im19.jpeg")
                                                    area3 = (776, 641,777,642)
                                                    im9_corte1 = im20.crop(area3)
                                                    im9_corte1.save('foto03.jpeg','jpeg')
                                                    im21 = Image.open('foto03.jpeg')
                                                    convert_im3 = im21.convert('RGB').getcolors()
                                                    print(convert_im3)

                                                    if convert_im3 == [(1, (255, 157, 40))] or convert_im3 == [(1, (255, 143, 33))] or convert_im3 == [(1, (255, 157, 68))] or convert_im3 == [(1, (252, 253, 255))]:
                                                        print("hello world")
                                                        time.sleep(2)

                                                        pyautogui.click((748, 651), interval= 0.5)

                                                        time.sleep(2)

                                                    pyautogui.hotkey('ctrl','f')
                                                    pyperclip.copy(text2)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    pyautogui.click((808, 651), interval=0.5)
                                                    
                                                    time.sleep(2)

                                                    pyautogui.click((883, 971), interval=0.5)

                                                    ##Adicionar metodo de espera de carregamento de janela
                                                    time.sleep(5)
                                                    a2 = 0
                                                    while a2 == 0:
                                                        ##adicionar metodo de espera de carregamento do site
                                                        im19 = ImageGrab.grab().save("im19.jpeg")
                                                        im20 = Image.open("im19.jpeg")
                                                        area3 = (416, 411, 417, 412)
                                                        im9_corte1 = im20.crop(area3)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im21 = Image.open('foto03.jpeg')
                                                        convert_im3 = im21.convert('RGB').getcolors()
                                                        print(convert_im3)

                                                        if convert_im3 == [(1, (240, 240, 240))]:
                                                            print("click on the bellow part 2")
                                                            pyautogui.click((1785, 180), interval = 1)
                                                            a2 = 1
                                                    
                                                    text3 = consulta[emp] + "_" + ano
                                                    pyperclip.copy(text3)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(1)
                                                    pyautogui.press('enter')
                                                    time.sleep(5)

                                                    pyautogui.hotkey('ctrl', 'w')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
                                                    
                                                    n = n + 1
                                                else:
                                                    time.sleep(3)
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
                                    path1 = os.listdir("E:\\Python_Project\\Data\\Downloads")
                                    for arquivo in path1:
                                        source = "E:\\Python_Project\\Data\\Downloads\\" + arquivo
                                        shutil.move(source, path)
                                    p = p +1
        
