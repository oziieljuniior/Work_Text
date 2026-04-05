from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import numpy as np
import time
import os
import shutil

import pyautogui
from PIL import ImageGrab, Image

# === CONFIGURAÇÕES ===
caminho_driver = '/home/darkcover1/Documentos/Work_Text/B3/WebDrivers/chromedriver-linux64_pc2/chromedriver'
caminho_excel = '/home/darkcover1/Documentos/Work_Text/B3/CNFB3.xlsx'
url_b3 = 'https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm'
text0 = 'Todos'
textod = ['Comentarios dos diretores', 'Comentario dos diretores']
textog = ['5. Gerenciamento de riscos e controles internos', '5. Politica de gerenciamento de riscos e controles internos']
text3 = 'Gerar PDF'

path_file1 = '/home/darkcover1/Documentos/Work_Text/B3/Códigos/Extractor/todos.png'

# === FUNÇÃO AUXILIAR ===
def abrir_pagina_inicial(driver, wait):
    driver.get(url_b3)
    time.sleep(20)
    try:
        botao_cookies = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
        )
        botao_cookies.click()
        print("🍪 Cookies aceitos.")
    except:
        print("✅ Nenhum banner de cookies detectado.")
    # Troca para o iframe de conteúdo
    iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(iframe)

# === EXECUÇÃO PRINCIPAL ===
df = pd.read_excel(caminho_excel)
empresas = df['TICKER'].dropna().tolist()

service = Service(executable_path=caminho_driver)
driver = webdriver.Chrome(service=service)

driver.maximize_window()

wait = WebDriverWait(driver, 20)

abrir_pagina_inicial(driver, wait)

for ticker in empresas:
    print(f"\n🔍 Buscando empresa: {ticker}")
    
    path = '/home/darkcover1/Documentos/Work_Text/B3/Códigos/Extractor/Data/'+ticker
    lista_empresas = os.listdir('/home/darkcover1/Documentos/Work_Text/B3/Códigos/Extractor/Data')
    catch = False
    for i in range(len(lista_empresas)):
        if lista_empresas[i] == ticker:
            print(lista_empresas[i], ticker)
            catch = True
            continue
    if catch is True:
        continue
    
    os.makedirs(path)
    relatorios_info = []

    try:
        # Preencher campo de busca
        campo = wait.until(EC.element_to_be_clickable((By.ID, "keyword")))
        campo.clear()
        campo.send_keys(ticker)
        time.sleep(1)

        # Clicar no botão "Buscar"
        botao = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Buscar") and not(@disabled)]')))
        botao.click()

        # Verificar se existe resultado
        try:
            resultado = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.card-body")))
            nome = resultado.find_element(By.CSS_SELECTOR, "h5.card-title2").text
            print(f"✅ Empresa encontrada: {ticker} → {nome}")
            resultado.click()

            # Aguarda a nova página da empresa carregar
            try:
                #####
                # Aguarda o select ficar disponível
                select_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'select[formcontrolname="selectMenu"]')))

                # Cria objeto do tipo Select
                select = Select(select_element)

                # Seleciona "Relatórios Estruturados" via valor (mais confiável)
                select.select_by_value("reports")

                # Alternativamente, poderia ser via texto visível:
                # select.select_by_visible_text("Relatórios Estruturados")

                print("📄 Relatórios Estruturados selecionado.")
                
                # Aguarda o select de anos
                select_ano = wait.until(EC.element_to_be_clickable((By.ID, "selectYear")))
                select_ano_obj = Select(select_ano)

                # Define o intervalo de anos
                anos_desejados = [str(ano) for ano in range(2010, 2025)]

                for ano in anos_desejados:
                    try:
                        select_ano_obj.select_by_value(ano)
                        print(f"📅 Ano selecionado: {ano}")
                        time.sleep(2)  # Aguarda carregar os links do ano

                        try:
                            # Verifica se o link do formulário está presente
                            link_ref = wait.until(
                                EC.presence_of_element_located((
                                    By.XPATH,
                                    f'//a[contains(text(), "{ano}") and contains(text(), "Formulário de Referência")]'
                                ))
                            )
                            href = link_ref.get_attribute("href")
                            print(f"📝 Formulário de Referência encontrado para {ano}: {href}")
                            name_ajustado = href.split('=')
                            save_me = ticker+"_"+name_ajustado[1]+"_"+str(ano) 
                            print(save_me)
                            relatorios_info.append(save_me)
                            

                            # Salva o identificador da aba atual
                            aba_original = driver.current_window_handle

                            # Clica no link (abre nova aba)
                            link_ref.click()

                            # Aguarda a nova aba abrir
                            wait.until(EC.number_of_windows_to_be(2))

                            # Alterna para a nova aba
                            abas = driver.window_handles
                            for aba in abas:
                                if aba != aba_original:
                                    driver.switch_to.window(aba)
                                    break

                            # Espera o botão "Salvar em PDF" aparecer
                            try:
                                botao_pdf = wait.until(EC.element_to_be_clickable((By.ID, "btnGeraRelatorioPDF")))
                                print("📄 Botão 'Salvar em PDF' localizado com sucesso.")
                                botao_pdf.click()  # ← só clique se realmente quiser baixar agora
                                ##########
                                # Espera o checkbox "Todos" da árvore jstree aparecer e clica para desmarcar                    
                                print('order 1')
                                #1) ctr + f -> 'Todos' -> enter
                                #2) Procura frame 1x1(em loop limitado(20 chamdas) -> clica na caixa
                                #3) Repete o processo para as outras frases
                                #4) Finaliza o processo chamando para baixar o pdf -> Salva em um caminho especifico
                                k1 = 0
                                while k1 <= 5:
                                    time.sleep(10)
                                    try:
                                        pyautogui.hotkey('ctrl','f')
                                        pyautogui.write(text0, interval = 0.05)
                                        #pyautogui.hotkey('ctrl', 'v')
                                        pyautogui.press('enter')
                                        print(f'Chamada: {text0} -> Screenshot + Procura')                                                        
                                        # Agora varre a tela procurando essa palavra
                                        im11 = Image.open('todos.png')
                                        guide = pyautogui.locateOnScreen(im11, confidence=0.9)
                                        # Verificando se a imagem foi encontrada
                                        if guide is not None:
                                            x = guide.left  # Posição X
                                            y = guide.top   # Posição Y
                                            click1 = (x,y)
                                            print(f"Posição X: {x}, Posição Y: {y}")
                                            pyautogui.click(click1)
                                            k1 = 6
                                            time.sleep(2)
                                            
                                            k2 = 0
                                            n1 = 0
                                            text1 = textod[n1]
                                            while k2 <= 5:
                                                time.sleep(3)
                                                try:
                                                    pyautogui.hotkey('ctrl','f')
                                                    print(text1)
                                                    pyautogui.write(text1, interval = 0.05)
                                                    #pyautogui.hotkey('ctrl', 'v')
                                                    pyautogui.press('enter')
                                                    time.sleep(2)

                                                    print(f'Chamada: {text1} -> Screenshot + Procura')                                                        
                                                    # Agora varre a tela procurando essa palavra
                                                    n2 = str(n1 + 1)
                                                    name1 = 'comentarios'+n2+'.png'
                                                    print(name1)
                                                    im12 = Image.open(name1)
                                                    guide = pyautogui.locateOnScreen(im12, confidence=0.9)
                                                    # Verificando se a imagem foi encontrada
                                                    if guide is not None:
                                                        x = guide.left  # Posição X
                                                        y = guide.top   # Posição Y
                                                        click1 = (x,y)
                                                        print(f"Posição X: {x}, Posição Y: {y}")
                                                        pyautogui.click(click1)
                                                        k2 = 6
                                                        time.sleep(2)

                                                        k3 = 0
                                                        n3 = 0
                                                        text2 = textog[n3]
                                                        while k3 <= 5:
                                                            time.sleep(3)
                                                            try:
                                                                pyautogui.hotkey('ctrl','f')
                                                                print(text2)
                                                                pyautogui.write(text2, interval = 0.05)
                                                                #pyautogui.hotkey('ctrl', 'v')
                                                                pyautogui.press('enter')
                                                                time.sleep(2)

                                                                print(f'Chamada: {text2} -> Screenshot + Procura')                                                        
                                                                # Agora varre a tela procurando essa palavra
                                                                n4 = str(n3 + 1)
                                                                name2 = 'gerenciamento'+n4+'.png'
                                                                print(name2)
                                                                im13 = Image.open(name2)
                                                                guide = pyautogui.locateOnScreen(im13, confidence=0.9)
                                                                # Verificando se a imagem foi encontrada
                                                                if guide is not None:
                                                                    x = guide.left  # Posição X
                                                                    y = guide.top   # Posição Y
                                                                    click1 = (x,y)
                                                                    print(f"Posição X: {x}, Posição Y: {y}")
                                                                    pyautogui.click(click1)
                                                                    k3 = 6

                                                                    pyautogui.click((1357, 686))
                                                                    time.sleep(2)
                                                                    pyautogui.click((701, 602))
                                                                    time.sleep(45)
                                                                    
                                                            except:
                                                                print(f'Botão "Gerenciamento" não encontrado"')
                                                                k3 += 1
                                                                if n3  > len(textog):
                                                                    n3 = 0
                                                                    text2 = textog[n3]
                                                                else:
                                                                    n3+=1
                                                                    text2 = textog[n3]
                                                                    
                                                except:
                                                    print(f'Botão "Comentario dos diretores" não encontrado"')
                                                    k2 += 1
                                                    if n1  > len(textod):
                                                        n1 = 0
                                                        text1 = textod[n1]
                                                    else:
                                                        n1+=1
                                                        text1 = textod[n1]
                                    except:
                                        print(f'Botão "Todos" não encontrado')   
                                        k1 += 1     
                                ##########

                                # Fechar a aba depois (opcional):
                                driver.close()
                                # driver.switch_to.window(aba_original)

                            except:
                                print("❌ Botão 'Salvar em PDF' não encontrado.")

                            # IMPORTANTE: Se quiser continuar o loop de outras empresas, volte à aba principal
                            driver.switch_to.window(aba_original)

                            # Reentrar no iframe após voltar da aba externa
                            try:
                                iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
                                driver.switch_to.frame(iframe)
                                print("↩️ Retornado com sucesso ao iframe.")
                            except:
                                print("⚠️ Não foi possível retornar ao iframe.")

                        except:
                            print(f"❌ Nenhum Formulário de Referência disponível para {ano}")

                    except Exception as e:
                        print(f"⚠️ Não foi possível selecionar o ano {ano}: {e}")
                ####

            except:
                print(f"⚠️ Página da empresa falhou para {ticker} — recarregando página inicial")
                driver.switch_to.default_content()
                abrir_pagina_inicial(driver, wait)

        except:
            alerta = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="alert" and contains(text(), "Não há dados disponíveis")]')))
            print(f"❌ Empresa não encontrada: {ticker}")
            time.sleep(2)

        lista_final = os.listdir('/home/darkcover1/Downloads/')
        #print(lista_final)
        print(len(lista_final))
        for i in range(len(lista_final)):
            name1 = lista_final[i].split("_")
            for j in range(len(relatorios_info)):
                name2 = relatorios_info[j].split("_")
                print(name1, name2)
                print(name1[0], name2[1])
                if name1[0] == name2[1]:
                    name_origem = '/home/darkcover1/Downloads/'+lista_final[i]
                    name_final =  path+"/"+str(name2[0])+'_'+str(name2[2])+'.pdf'
                    print(f'nome original: {name_origem} \nnome final: {name_final}')
                    shutil.move(name_origem, name_final)
                

    except Exception as e:
        print(f"❌ Erro ao processar {ticker}: {e}")
        driver.switch_to.default_content()
        abrir_pagina_inicial(driver, wait)

driver.quit()
