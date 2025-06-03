from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import time
import pyautogui
from PIL import Image

# === CONFIGURAÇÕES ===
caminho_driver = '/home/darkcover1/Documentos/Work_Text/B3/WebDrivers/chromedriver-linux64_pc2/chromedriver'
caminho_excel = '/home/darkcover1/Documentos/Work_Text/B3/CNFB3.xlsx'
url_b3 = 'https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm'

text0 = 'Todos'
textod = ['Comentarios dos diretores', 'Comentario dos diretores']
text2 = '5. Gerenciamento de riscos e controles internos'

# === FUNÇÃO AUXILIAR ===
def abrir_pagina_inicial(driver, wait):
    driver.get(url_b3)
    time.sleep(2)
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
        print("🍪 Cookies aceitos.")
    except:
        print("✅ Nenhum banner de cookies detectado.")
    iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(iframe)

# === EXECUÇÃO PRINCIPAL ===
df = pd.read_excel(caminho_excel)
tickers = df['TICKER'].dropna().tolist()

service = Service(executable_path=caminho_driver)
driver = webdriver.Chrome(service=service)
driver.maximize_window()
wait = WebDriverWait(driver, 20)

abrir_pagina_inicial(driver, wait)

for ticker in tickers:
    print(f"\n🔍 Buscando empresa: {ticker}")
    try:
        campo = wait.until(EC.element_to_be_clickable((By.ID, "keyword")))
        campo.clear()
        campo.send_keys(ticker)
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Buscar") and not(@disabled)]'))).click()

        try:
            resultado = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.card-body")))
            nome = resultado.find_element(By.CSS_SELECTOR, "h5.card-title2").text
            print(f"✅ Empresa encontrada: {ticker} → {nome}")
            resultado.click()

            try:
                select_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'select[formcontrolname="selectMenu"]')))
                Select(select_element).select_by_value("reports")
                print("📄 Relatórios Estruturados selecionado.")

                select_ano = wait.until(EC.element_to_be_clickable((By.ID, "selectYear")))
                select_ano_obj = Select(select_ano)
                anos_desejados = [str(ano) for ano in range(2020, 2025)]

                for ano in anos_desejados:
                    try:
                        select_ano_obj.select_by_value(ano)
                        print(f"📅 Ano selecionado: {ano}")
                        time.sleep(2)

                        try:
                            link_ref = wait.until(
                                EC.presence_of_element_located((
                                    By.XPATH,
                                    f'//a[contains(text(), "{ano}") and contains(text(), "Formulário de Referência")]'
                                ))
                            )
                            link_ref.click()
                            wait.until(EC.number_of_windows_to_be(2))
                            nova_aba = [a for a in driver.window_handles if a != driver.current_window_handle][0]
                            driver.switch_to.window(nova_aba)

                            try:
                                wait.until(EC.element_to_be_clickable((By.ID, "btnGeraRelatorioPDF"))).click()
                                time.sleep(5)

                                def localizar_e_clicar(termo, imagem):
                                    pyautogui.hotkey('ctrl', 'f')
                                    pyautogui.write(termo, interval=0.05)
                                    pyautogui.press('enter')
                                    for _ in range(5):
                                        pos = pyautogui.locateOnScreen(imagem, confidence=0.9)
                                        if pos:
                                            pyautogui.click(pos.left, pos.top)
                                            return True
                                        time.sleep(1)
                                    return False

                                if localizar_e_clicar(text0, 'todos.png'):
                                    for i, texto in enumerate(textod):
                                        nome_img = f'comentarios{i+1}.png'
                                        if localizar_e_clicar(texto, nome_img):
                                            if localizar_e_clicar(text2, 'gerenciamento.png'):
                                                break
                            
                            except:
                                print("❌ Botão 'Salvar em PDF' não encontrado.")

                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            driver.switch_to.default_content()
                            abrir_pagina_inicial(driver, wait)

                        except:
                            print(f"❌ Nenhum Formulário de Referência disponível para {ano}")

                    except Exception as e:
                        print(f"⚠️ Erro ao selecionar o ano {ano}: {e}")

            except:
                print(f"⚠️ Falha na página da empresa {ticker}, recarregando...")
                driver.switch_to.default_content()
                abrir_pagina_inicial(driver, wait)

        except:
            alerta = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="alert" and contains(text(), "Não há dados disponíveis")]')))
            print(f"❌ Empresa não encontrada: {ticker}")
            time.sleep(2)

    except Exception as e:
        print(f"❌ Erro ao processar {ticker}: {e}")
        driver.switch_to.default_content()
        abrir_pagina_inicial(driver, wait)


driver.quit()
