from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import time

# === CONFIGURAÇÕES ===
caminho_driver = '/home/darkcover1/Documentos/Work_Text/B3/WebDrivers/chromedriver-linux64_pc2/chromedriver'
caminho_excel = '/home/darkcover1/Documentos/Work_Text/B3/CNFB3.xlsx'
url_b3 = 'https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm'

# === FUNÇÃO AUXILIAR ===
def abrir_pagina_inicial(driver, wait):
    driver.get(url_b3)
    time.sleep(2)
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
wait = WebDriverWait(driver, 20)

abrir_pagina_inicial(driver, wait)

for ticker in empresas:
    print(f"\n🔍 Buscando empresa: {ticker}")

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
                                time.sleep(5)
                                # Espera o checkbox "Todos" da árvore jstree aparecer e clica para desmarcar
                                ##
                                # Aguarda o iframe do modal abrir
                                try:
                                    iframe_modal = wait.until(EC.presence_of_element_located((By.ID, "iFrameModal")))
                                    driver.switch_to.frame(iframe_modal)
                                    print("🧭 Entrou no iframe do modal com sucesso.")
                                    time.sleep(2)
                                    # Clica no checkbox da raiz (Todos)
                                    try:
                                        checkbox_todos = wait.until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, ".jstree-anchor > .jstree-checkbox"))
                                        )
                                        checkbox_todos.click()
                                        print("☑️ Checkbox 'Todos' desmarcado dentro do modal.")
                                        time.sleep(1)
                                    except Exception as e:
                                        print(f"❌ Não foi possível desmarcar 'Todos' no iframe: {e}")
                                    # Voltar ao contexto principal da aba (fora do iframe)
                                    driver.switch_to.default_content()

                                except Exception as e:
                                    print(f"❌ Erro ao acessar o iframe do modal: {e}")
                                ##

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

    except Exception as e:
        print(f"❌ Erro ao processar {ticker}: {e}")
        driver.switch_to.default_content()
        abrir_pagina_inicial(driver, wait)

driver.quit()
