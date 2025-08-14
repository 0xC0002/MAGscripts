from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time

# === INPUTS ===
competencia = input("Digite a competência (AAAA/MM): ")
fornecedor = input("Digite o fornecedor (igual ao site): ")
colaborador = input("Digite o colaborador (igual ao site): ")
horas = input("Digite a quantidade de horas (ex: 8,00): ")
pdf_path = input("Digite o caminho completo do PDF: ")

# === CONFIGURAÇÃO DO EDGE ===
edge_service = EdgeService(executable_path="msedgedriver.exe")
driver = webdriver.Edge(service=edge_service)
driver.get("http://governanca.mongeral.seguros")

wait = WebDriverWait(driver, 30)

# === CLICAR EM "ADICIONAR NOVO ITEM" ===
add_item = wait.until(EC.element_to_be_clickable((By.ID, "idHomePageNewItem")))
add_item.click()

# === PREENCHER COMPETÊNCIA ===
campo_competencia = wait.until(EC.presence_of_element_located((
    By.XPATH, "//input[contains(@title, 'Formato AAAAMM')]"
)))
campo_competencia.send_keys(competencia)

# === SELECIONAR FORNECEDOR ===
select_fornecedor_elem = wait.until(EC.presence_of_element_located((
    By.XPATH, "//select[contains(@name, 'lookup') and contains(@class, 'nf-associated-control')]"
)))
select_fornecedor = Select(select_fornecedor_elem)
select_fornecedor.select_by_visible_text(fornecedor)

# === SELECIONAR COLABORADOR ===
select_colaborador_elem = wait.until(EC.presence_of_element_located((
    By.XPATH, "//select[contains(@name, 'lookup') and contains(@class, 'nf-associated-control')][2]"
)))
select_colaborador = Select(select_colaborador_elem)
select_colaborador.select_by_visible_text(colaborador)

# === MARCAR RADIO "DECIMAL" ===
radio_decimal = wait.until(EC.presence_of_element_located((
    By.XPATH, "//input[@type='radio' and @value='Decimal']"
)))
radio_decimal.click()

# === PREENCHER QUANTIDADE DE HORAS ===
campo_horas = wait.until(EC.presence_of_element_located((
    By.XPATH, "//input[contains(@class, 'txtQuantDecimal')]"
)))
campo_horas.clear()
campo_horas.send_keys(horas)

print("✅ Formulário preenchido com sucesso! Verifique no Edge antes de enviar.")
time.sleep(20)  # tempo para revisar

driver.quit()
