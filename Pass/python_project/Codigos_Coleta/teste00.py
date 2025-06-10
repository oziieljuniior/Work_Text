import pyautogui
import time
from PIL import Image
import pandas as pd
import pyscreenshot as ImageGrab
import os
import shutil

data = pd.read_excel('/home/oziel/Documentos/Alunos/PauloB/data/Cias não financeiras - 12-12-2022.xlsx')
consulta = data['ticker'].to_list()
t1 = len(consulta)
for emp in range(170,t1):
    path = "/home/oziel/Documentos/Alunos/PauloB/data/empresas/" + consulta[emp]
    os.makedirs(path)
    
    i = 0
    while i == 0:
        print("Fase 1")
        im1 = ImageGrab.grab().save("im1.png")
        im2 = Image.open("../im1.png")
        #[(1, (251, 251, 251))] [(1, (249, 249, 251))]
        area1 = (200,47,201,48)
        area2 = (188,208,189,209)
        im3_corte1 = im2.crop(area1)
        im4_corte2 = im2.crop(area2)
        im3_corte1.save('foto01.jpeg','jpeg')
        im4_corte2.save('foto02.jpeg','jpeg')
        im5 = Image.open('../foto01.jpeg')
        im6 = Image.open('../foto02.jpeg')
        convert_im1 = im5.convert('RGB').getcolors()
        convert_im2 = im6.convert('RGB').getcolors()
        print(convert_im1,convert_im2)
        
        if convert_im1 == [(1, (251, 251, 251))] and convert_im2 == [(1, (249, 249, 251))]:
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
                im8 = Image.open("../im2.png")
                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                area3 = (136,289,137,290)
                area4 = (323,679,324,680)
                im9_corte1 = im8.crop(area3)
                im10_corte2 = im8.crop(area4)
                im9_corte1.save('foto03.jpeg','jpeg')
                im10_corte2.save('foto04.jpeg','jpeg')
                im11 = Image.open('../foto03.jpeg')
                im12 = Image.open('../foto04.jpeg')
                convert_im3 = im11.convert('RGB').getcolors()
                convert_im4 = im12.convert('RGB').getcolors()
                print(convert_im3,convert_im4)
                if convert_im3 == [(1, (1, 44, 99))] and convert_im4 == [(1, (247, 247, 247))]:
                    print("Fase 4")
                    k = 0
                    time.sleep(10)
                    pyautogui.doubleClick(410,676)
                    pyautogui.write(consulta[emp], interval = 0.5)
                    pyautogui.press('enter')
# =============================================================================
#                     time.sleep(5)
# =============================================================================
                    while k == 0:
                        print("Fase 5")
                        
                        im13 = ImageGrab.grab().save("im3.jpeg")
                        im14 = Image.open("../im3.jpeg")
                        #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                        area5 = (105,309,106,310)
                        area6 = (252,659,253,660)
                        im15_corte1 = im14.crop(area5)
                        im16_corte2 = im14.crop(area6)
                        im15_corte1.save('foto05.jpeg','jpeg')
                        im16_corte2.save('foto06.jpeg','jpeg')
                        im17 = Image.open('../foto05.jpeg')
                        im18 = Image.open('../foto06.jpeg')
                        convert_im5 = im11.convert('RGB').getcolors()
                        convert_im6 = im12.convert('RGB').getcolors()
                        print(convert_im5,convert_im6)
                        if convert_im3 == [(1, (1, 44, 99))] and convert_im4 == [(1, (247, 247, 247))]:
                            time.sleep(3)
                            l = 0
                            print("Fase 6")
                            time.sleep(10)
                            pyautogui.doubleClick(511,685)
                            #possivel colocar outra captura
                            time.sleep(3)
                            pyautogui.click(1503,517)
                            time.sleep(3)
                            pyautogui.click(1490,586)
                            while l == 0:
                                print("Fase 7")
                                im13 = ImageGrab.grab().save("im3.jpeg")
                                im14 = Image.open("../im3.jpeg")
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                #[(1, (1, 44, 99))] [(1, (247, 247, 247))]
                                area5 = (105,309,106,310)
                                area6 = (252,659,253,660)
                                im15_corte1 = im14.crop(area5)
                                im16_corte2 = im14.crop(area6)
                                im15_corte1.save('foto05.jpeg','jpeg')
                                im16_corte2.save('foto06.jpeg','jpeg')
                                im17 = Image.open('../foto05.jpeg')
                                im18 = Image.open('../foto06.jpeg')
                                convert_im5 = im11.convert('RGB').getcolors()
                                convert_im6 = im12.convert('RGB').getcolors()
                                print(convert_im5,convert_im6)
                                if convert_im3 == [(1, (1, 44, 99))] and convert_im4 == [(1, (247, 247, 247))]:
                                    print("Fase 8")
                                    time.sleep(10)
                                    pyautogui.click(477,591)
                                    #
                                    
                                    
                                    m = 0
                                    while m <= 11:
                                        cot = 2010 + m
                                        ano = str(cot)
                                        if cot == 2010:
                                            n = 0
                                            print("Ano ", ano)
                                            time.sleep(3)
                                            pyautogui.click(431,937)                                            
                                            time.sleep(3)
                                            #pyautogui.click((1897,89),interval = 1)
                                            #time.sleep(1)
                                            #pyautogui.click((1699,524),interval = 1)
                                            #pyautogui.press('enter')
                                            
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao ')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                #[(1, (236, 76, 44))] [(1, (232, 85, 33))]
                                                #[(1, (71, 209, 124))] [(1, (54, 213, 121))]
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                                                                        
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            time.sleep(3)
                                            pyautogui.click(437,910)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            time.sleep(3)
                                            pyautogui.click(457,886)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                        
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(452,863)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o= o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(449,834)
                                            
                                            time.sleep(3)
                                            
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(433,811)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(458,787)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(454,758)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            
                                            time.sleep(3)
                                            pyautogui.click(466,737)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            time.sleep(3)
                                            pyautogui.click(455,712)
                                            
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            time.sleep(3)
                                            pyautogui.click(454,687)
                                            
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    time.sleep(2)
                                                    pyautogui.click((79,683), interval = 2)
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
                                            pyautogui.click((1901,1060), interval = 0.5)
                                            
                                            time.sleep(3)
                                            pyautogui.click(477,591)
                                            time.sleep(3)
                                            pyautogui.click(467,663)
                                            
                                            time.sleep(3)
                                            pyautogui.hotkey('ctrl','f')
                                            pyautogui.write('Demonstracoes Financeiras Padronizadas - Versao')
                                            pyautogui.press('enter')
                                            while n == 0:
                                                im7 = ImageGrab.grab().save("im2.jpeg")
                                                im8 = Image.open("../im2.jpeg")
                                                #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                area3 = (627,592,628,593)
                                                area4 = (830,596,831,597)
                                                im9_corte1 = im8.crop(area3)
                                                im10_corte2 = im8.crop(area4)
                                                im9_corte1.save('foto03.jpeg','jpeg')
                                                im10_corte2.save('foto04.jpeg','jpeg')
                                                im11 = Image.open('../foto03.jpeg')
                                                im12 = Image.open('../foto04.jpeg')
                                                convert_im3 = im11.convert('RGB').getcolors()
                                                convert_im4 = im12.convert('RGB').getcolors()
                                                print(convert_im3,convert_im4)
                                                #print(convert_im1)
                                                if ((convert_im3 == [(1, (236, 76, 44))] and (convert_im4 == [(1, (232, 85, 33))])) or (((convert_im3 == [(1, (71, 209, 124))] and (convert_im4 ==  [(1, (54, 213, 121))]))))):
                                                    print("encontrei ouro")
                                                    time.sleep(5)
                                                    
                                                    pyautogui.click((627,592), interval = 1)

                                                    time.sleep(15)
                                                    pyautogui.click((1659,199), interval = 1)
                                                    time.sleep(3)
                                                    pyautogui.click((677,461), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((677,621), interval = 1)
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    pyautogui.press('pagedown')
                                                    time.sleep(2)
                                                    pyautogui.click((703,631), interval = 1)
                                                    time.sleep(2)
                                                    pyautogui.click((936,934), interval = 1)
                                                    time.sleep(5)
                                                    o = 0

                                                    while o == 0:
                                                        im7 = ImageGrab.grab().save("im2.jpeg")
                                                        im8 = Image.open("../im2.jpeg")
                                                        #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
                                                        area3 = (491,243,492,244)
                                                        area4 = (352,235,353,236)
                                                        im9_corte1 = im8.crop(area3)
                                                        im10_corte2 = im8.crop(area4)
                                                        im9_corte1.save('foto03.jpeg','jpeg')
                                                        im10_corte2.save('foto04.jpeg','jpeg')
                                                        im11 = Image.open('../foto03.jpeg')
                                                        im12 = Image.open('../foto04.jpeg')
                                                        convert_im3 = im11.convert('RGB').getcolors()
                                                        convert_im4 = im12.convert('RGB').getcolors()
                                                        print(convert_im3,convert_im4)
                                                        #print(convert_im1)
                                                        #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
                                                        if convert_im3 == [(1, (255, 255, 255))] and convert_im4 == [(1, (237, 237, 239))]:
                                                            print("hello world")
                                                            pyautogui.click((927,50), interval = 2)
                                                            pyautogui.click((702,50), interval = 2)
                                                            o = o + 1
                                                    
                                                    
                                                    pyautogui.press('pageup')
                                                    pyautogui.press('pageup')
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
                                    path1 = os.listdir("/home/oziel/Downloads")
                                    for arquivo in path1:
                                        source = "/home/oziel/Downloads/" + arquivo
                                        shutil.move(source, path)
                                    p = p +1
        