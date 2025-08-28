from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import time
import csv
import os
import subprocess

# === Carregar variáveis de ambiente ===
load_dotenv()
LOGIN = os.getenv("CRM_LOGIN")
SENHA = os.getenv("CRM_SENHA")
CRM_URL = os.getenv("CRM_URL")

# === Configuração do navegador ===
options = Options()
options.add_argument('--start-maximized')
options.add_argument('--disable-blink-features=AutomationControlled')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# === Acessar o site e fazer login ===
driver.get(CRM_URL)
time.sleep(3)

driver.find_element(By.NAME, "login").send_keys(LOGIN)
driver.find_element(By.NAME, "senha").send_keys(SENHA)
driver.find_element(By.CSS_SELECTOR, 'input[type="submit"]').click()
time.sleep(5)

# === Navegar até CRM Imóveis > Imóveis: Listar ===
driver.find_element(By.XPATH, "//span[contains(text(), 'CRM Imóveis')]").click()
time.sleep(1)
driver.find_element(By.XPATH, "//a[contains(text(), 'Imóveis: Listar')]").click()
time.sleep(4)

# === Pesquisa por PSPP ===
driver.find_element(By.ID, "pesquisa").clear()
driver.find_element(By.ID, "pesquisa").send_keys("PSPP")
driver.find_element(By.NAME, "btngv2").click()
time.sleep(4)

# === Criar arquivo CSV ===
with open('dados_dos_imoveis.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(["Código", "Título", "Condomínio", "Referência", "Status / Datas"])

# === Função para extrair dados dos detalhes ===
def extrair_detalhes(link):
    driver.execute_script("window.open(arguments[0]);", link)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(4)

    try:
        titulo = driver.find_element(By.TAG_NAME, "h3").text
        condominio = ""
        paragrafos = driver.find_elements(By.TAG_NAME, "p")
        for p in paragrafos:
            if "Condomínio" in p.text:
                condominio = p.text.split("Condomínio: ")[-1].split(" - ")[0]
                break

        referencia = driver.find_element(By.XPATH, "//strong[contains(text(),'Referência')]").text.split(":")[-1].strip()
        status_data = driver.find_element(By.XPATH, "//p[contains(text(),'Cadastrado')]").text
        codigo = link.split("cod=")[-1].split("&")[0]

        with open('dados_dos_imoveis.csv', 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([codigo, titulo, condominio, referencia, status_data])

        print(f"✅ Imóvel {codigo} extraído com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao extrair detalhes: {e}")

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(2)

# === Loop nas páginas ===
pagina_atual = 1
limite_paginas = 4

while pagina_atual <= limite_paginas:
    print(f"\n🔎 Página {pagina_atual}...\n")
    time.sleep(2)

    botoes_detalhes = driver.find_elements(By.XPATH, '//a[@title="Detalhes" and contains(@href,"imovel_vida_frame")]')
    print(f"➡️ {len(botoes_detalhes)} imóveis encontrados.")

    for botao in botoes_detalhes:
        link_detalhe = botao.get_attribute("href")
        extrair_detalhes(link_detalhe)

    pagina_atual += 1
    try:
        proxima = driver.find_element(By.XPATH, f'//a[contains(@href, "pag={pagina_atual}")]')
        proxima.click()
        time.sleep(5)
    except:
        print("⚠️ Sem próxima página disponível.")
        break

# === Finalização ===
print("\n🏁 Extração finalizada com sucesso!")
driver.quit()

# === Executa o próximo script ===
print("\n✅ Iniciando geração dos relatórios...\n")
subprocess.run(["python", "Pesquisa2.py"])
