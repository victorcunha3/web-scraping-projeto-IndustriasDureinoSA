import os
import re
import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from typing import List

def main():
    #chromedriver = r"C:\Users\Aprendiz\Downloads\chromedriver.exe"#driver local
    chromedriver = r"c:\Users\Aprendiz\Downloads\chromedriver-win64\chromedriver.exe"
    service = Service(chromedriver)
    driver = webdriver.Chrome(service=service)
    driver.get('https://xpert.com.br/atg/')
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')

    if len(iframes) > 0:
        driver.switch_to.frame(iframes[0])

    wait = WebDriverWait(driver, 20)
    username_field = wait.until(EC.presence_of_element_located((By.ID, 'inputUser')))
    password_field = wait.until(EC.presence_of_element_located((By.ID, 'inputPass')))

    username_field.send_keys('')
    password_field.send_keys('xpert')

    form = driver.find_element(By.TAG_NAME, 'form')
    form.submit()

    tanques = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.content.border-solid')))

    dados_tanques:List = []

    for tanque in tanques:
        capacidade_element = tanque.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
        litros_element = tanque.find_element(By.CSS_SELECTOR, 'h1.leitura-tanque > b')
        porcentagem_element = tanque.find_element(By.CSS_SELECTOR, 'span.ng-binding')
        h3_element = tanque.find_element(By.XPATH, './/h3')
        h3_text = h3_element.text.strip()

        capacidade_text = capacidade_element.text
        capacidade_match = re.search(r'[\d.]+', capacidade_text)
        capacidade = capacidade_match.group() if capacidade_match else ''

        litros_text = litros_element.text
        litros_match = re.search(r'(\d+(\.\d+)?)', litros_text)
        litros = litros_match.group(1) if litros_match else '0.0'

        porcentagem = porcentagem_element.text

        data_execucao = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        dados = {
            "Data Execução": data_execucao,
            "Litros": litros,
            "Capacidade": capacidade,
            "Porcentagem": porcentagem,
            "Nome": h3_text
        }
        dados_tanques.append(dados)

    driver.quit()

    file_path = r"t:\FABRICA\PRODUCAO\Compartilhado\4 - SETORES\4.1 - Ensacado\hexanoTeste.xlsx"

    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path)
        existing_data = existing_df.to_dict(orient='records')
        dados_tanques = existing_data + dados_tanques

    try:
        if len(dados_tanques) > 0:
            df = pd.DataFrame(dados_tanques)
            df.to_excel(file_path, index=False)
            print("Os dados foram exportados para o Excel com sucesso.")
        else:
            print("Não há dados para exportar.")
    except PermissionError:
        print("ERRO! EXCEL ABERTO! Feche o arquivo e tente novamente.")

if __name__ == "__main__":
    main()
