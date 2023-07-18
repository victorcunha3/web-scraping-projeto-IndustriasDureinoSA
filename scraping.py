from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import pandas as pd

def buscar_tanques() -> list:
    chromedriver = r"C:\Users\Aprendiz\Downloads\chromedriver.exe"
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
    password_field.send_keys('')

    form = driver.find_element(By.TAG_NAME, 'form')
    form.submit()

    tanques = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.tanque')))

    dados_tanques = []

    for tanque in tanques:
        #informacao_tanque = tanque.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
        capacidade_element = tanque.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
        litros_element = tanque.find_element(By.CSS_SELECTOR, 'h1.leitura-tanque > b')
        porcentagem_element = tanque.find_element(By.CSS_SELECTOR, 'span.ng-binding')
        #tanque_element = tanque.find_element(By.CSS_SELECTOR, 'div.tanque-label > h3')
        #tanque_element = tanque.find_element(By.XPATH, './/h3[contains(text(), "tanque")]')
        #tanque_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'h3.ng-binding')))
        #tanque = tanque_element.text

        capacidade_text = capacidade_element.text
        capacidade_match = re.search(r'[\d.]+', capacidade_text)
        capacidade = capacidade_match.group() if capacidade_match else ''

        #tanque_nome = tanque_element.text
        #tanque_match = re.search(r'(\d+)\s*-\s*tanque\s*\d+\s*(\w+)', tanque_nome)
        #tanque_info = f'{tanque_match.group(1)} - tanque {tanque_match.group(2)}' if tanque_match else ''

        litros_text = litros_element.text
        litros_match = re.search(r'[\d.]+', litros_text)
        litros = litros_match.group() if litros_match else ''
        porcentagem = porcentagem_element.text

        dados = {
            "Litros": litros,
            "Capacidade": capacidade,
            "porcentagem": porcentagem,
            #"Tanque": tanque_element
        }

        dados_tanques.append(dados)

    print(dados_tanques)

    driver.quit()


import pandas as pd

dados_tanques = buscar_tanques()
df = pd.DataFrame(dados_tanques)
df.to_excel('dados_tanques.xlsx', index=False)
