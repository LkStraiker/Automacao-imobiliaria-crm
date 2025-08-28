from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import time
import csv
import os
from datetime import datetime
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

# === Função para salvar log ===
def salvar_log(codigo, valor_antigo, valor_novo):
    with open('log.csv', 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([codigo, valor_antigo, valor_novo, datetime.now().strftime("%d/%m/%Y %H:%M:%S")])

# === Função para ir para a página N ===
def ir_para_pagina(numero):
    try:
        botao = driver.find_element(By.XPATH, f'//a[contains(@href, "pag={numero}") and contains(@class, "lipagina-btn-paginacao")]')
        botao.click()
        print(f"➡️ Indo para a página {numero}...")
        time.sleep(4)
        return True
    except:
        print(f"❌ Página {numero} não encontrada.")
        return False

# === Acesso ao sistema ===
driver.get(CRM_URL)
time.sleep(3)

driver.find_element(By.NAME, "login").send_keys(LOGIN)
driver.find_element(By.NAME, "senha").send_keys(SENHA)
driver.find_element(By.CSS_SELECTOR, 'input[type="submit"]').click()
time.sleep(5)

# === Navegar até área de imóveis ===
driver.find_element(By.XPATH, "//span[contains(text(), 'CRM Imóveis')]").click()
time.sleep(1)
driver.find_element(By.XPATH, "//a[contains(text(), 'Imóveis: Listar')]").click()
time.sleep(5)

# === Pesquisa por código PSPP ===
select_tipo = Select(driver.find_element(By.NAME, "tipo"))
select_tipo.select_by_value("")

pesquisa_input = driver.find_element(By.ID, "pesquisa")
pesquisa_input.clear()
pesquisa_input.send_keys("PSPP")
driver.find_element(By.NAME, "btngv2").click()
time.sleep(5)

# === Criar log de alterações ===
with open('log.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(["Código", "Valor IPTU (Antes)", "Valor IPTU (Depois)", "Data e Hora"])

# === Loop nas páginas ===
for pagina in range(1, 5):
    print(f"\n📄 Processando página {pagina}...\n")
    time.sleep(2)

    links_alterar = driver.find_elements(By.XPATH, '//a[contains(@href, "imovel_alterar.php?cod=") and @title="Alterar"]')
    print(f"➡️ {len(links_alterar)} anúncios encontrados na página {pagina}.")

    for i, link in enumerate(links_alterar, start=1):
        href = link.get_attribute("href")
        codigo = href.split("cod=")[-1]
        print(f"🔧 [{i}] Editando anúncio {codigo}...")

        driver.execute_script("window.open(arguments[0]);", href)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(4)

        try:
            iptu_input = driver.find_element(By.ID, "valor_iptu")
            valor_atual = iptu_input.get_attribute("value").replace(".", "").replace(",", ".").strip()

            if valor_atual:
                valor_original = float(valor_atual)
                novo_valor = valor_original + 2
            else:
                valor_original = 0
                novo_valor = 2.00

            iptu_input.clear()
            iptu_input.send_keys(f"{novo_valor:.2f}".replace(".", ","))

            driver.find_element(By.CSS_SELECTOR, 'input.btnSalvarImovel').click()
            time.sleep(4)

            salvar_log(codigo, f"{valor_original:.2f}".replace(".", ","), f"{novo_valor:.2f}".replace(".", ","))
            print(f"✅ Anúncio {codigo} salvo com sucesso.")

        except Exception as e:
            print(f"⚠️ Erro ao editar imóvel {codigo}: {e}")

        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(2)

    if not ir_para_pagina(pagina + 1):
        break

print("🏁 Fim da automação.")
driver.quit()


