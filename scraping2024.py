import os
import re
import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import typing
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from openpyxl import load_workbook, Workbook

driver:str = r"c:\Users\aprendiz.DUREINO\Downloads\driverCurrent\chromedriver-win64\chromedriver.exe"

def iniciar_navegador(url:str):
    global driver
    iniciar_chromedriver = Service(driver)
    driver = webdriver.Chrome(service=iniciar_chromedriver)
    print(driver)
    driver.get(url)
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    if len(iframes) > 0:
        driver.switch_to.frame(iframes[0])

def buscar_campos(usuarioID:str, senhaID:str):
    wait = WebDriverWait(driver, 20)
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, usuarioID))) #inputUser
    campo_senha = wait.until(EC.presence_of_element_located((By.ID, senhaID))) #inputPass
    campo_usuario.send_keys('5169')
    campo_senha.send_keys('xpert')
    time.sleep(2)
    formulario = driver.find_element(By.TAG_NAME, 'form')
    print(formulario)
    formulario.submit()

def busca_tanques():
    wait = WebDriverWait(driver, 10)
    try:
        tanques = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.content.border-solid')))
        print(tanques)
    except:
        print("-> Tanques não foram encontrados!!!!")
    dados_tanques: typing.List = []

    time.sleep(6)


    """for tanqueAtual in dados_tanques:
        capacidade_element = tanqueAtual.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
        litros_element = tanqueAtual.find_element(By.CSS_SELECTOR, 'h1.leitura-tanque > b')
        porcentagem_element = tanqueAtual.find_element(By.CSS_SELECTOR, 'span.ng-binding')
        h3_element = tanqueAtual.find_element(By.XPATH, './/h3')
        h3_text = h3_element.text.strip()
        print(litros_element)

        capacidade_text = capacidade_element.text
        litros_text = litros_element.text
        #litros = (re.sub(r'[^\d.]', '', litros_text)) if litros_text else 0.0
        porcentagem_text = porcentagem_element.text

        data_execucao = datetime.datetime.now().strftime('%d-%m-%Y')
        hora_execucao = datetime.datetime.now().strftime('%H:%M:%S')

        dados = {
            "Data Execução": data_execucao,
            "Litros": litros_text,
            "Capacidade": capacidade_text,
            "Porcentagem": porcentagem_text,
            "Nome": h3_text,
            "Hora Execução": hora_execucao
        }
        dados_tanques.append(dados)
        print(dados)"""


def main():
    iniciar_navegador("https://xpert.com.br/atg/")
    buscar_campos("inputUser","inputPass")
    busca_tanques()
    time.sleep(5)


if __name__ == "__main__":
    main()

