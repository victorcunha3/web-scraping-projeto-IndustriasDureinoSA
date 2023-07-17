from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re

chromedriver = r"C:\Users\victo\Downloads\chromedriver_win32\chromedriver.exe"

service = Service(chromedriver)

driver = webdriver.Chrome(service=service)

driver.get('https://xpert.com.br/atg/')

iframes = driver.find_elements(By.TAG_NAME, 'iframe')

if len(iframes) > 0:
    driver.switch_to.frame(iframes[0])

wait = WebDriverWait(driver, 20)
username_field = wait.until(EC.presence_of_element_located((By.ID, 'inputUser')))
password_field = wait.until(EC.presence_of_element_located((By.ID, 'inputPass')))

username_field.send_keys('5169')
password_field.send_keys('no password')

form = driver.find_element(By.TAG_NAME, 'form')
form.submit()

informacao_tanque = wait.until(EC.presence_of_element_located((By.XPATH, '//h4[contains(text(), "Capacidade")]/b')))
#capacidade_element = wait.until(EC.presence_of_element_located((By.XPATH, '//h4[contains(text(), "Capacidade")]/b')))
capacidade_element = wait.until(EC.presence_of_element_located((By.XPATH, '//h4[contains(text(), "Capacidade")]/b')))
#litros_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'b.ng-binding')))
litros_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'h1.leitura-tanque > b')))
porcentagem_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.ng-binding')))

tanque_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'h3.ng-binding')))

#capacidade = capacidade_element.text
# Extrair os dados
capacidade_text = capacidade_element.text
capacidade_match = re.search(r'[\d.]+', capacidade_text)
capacidade = capacidade_match.group() if capacidade_match else ''
#litros = litros_element.text

tanque = tanque_element.text

# Imprimir os dados
print("Tanque:", tanque)

litros_text = litros_element.text
litros_match = re.search(r'[\d.]+', litros_text)
litros = litros_match.group() if litros_match else ''
porcentagem = porcentagem_element.text

#To do
def extrair_dados():
    ...

print(f"Capacidade: {capacidade}")
print(f"Litros: {litros}")
print(f"Porcentagem: {porcentagem}")

driver.quit()
