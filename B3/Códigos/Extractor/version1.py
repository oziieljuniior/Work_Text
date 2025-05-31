from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

# Caminho para o ChromeDriver compatível
caminho_driver = '/home/darkcover/Documentos/B3/WebDrivers/chromedriver-linux64/chromedriver'
# Inicializa o driver

service = Service(executable_path=caminho_driver)
driver = webdriver.Chrome(service=service)

# Acessa o site da B3
driver.get("https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm")

# Espera para garantir carregamento da página
time.sleep(5)

# Encerra o navegador (após seus comandos)
driver.quit()
