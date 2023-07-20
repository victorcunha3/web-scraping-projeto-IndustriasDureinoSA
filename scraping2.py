import os
import csv
import re
import pandas as pd
import datetime
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog, QTableWidgetItem, QTableWidget, QHeaderView
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from qt_material import apply_stylesheet, QtStyleTools, QUiLoader


class MainWindow(QMainWindow, QtStyleTools):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Busca de Tanques(desenvolvido por Victor)")
        self.setMinimumSize(500, 400)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.label = QLabel("Dados dos Tanques:")
        self.layout.addWidget(self.label)

        self.table = QTableWidget()
        self.table.setColumnCount(4) 
        self.table.setHorizontalHeaderLabels(["Data Execução", "Litros", "Capacidade", "Porcentagem"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.layout.addWidget(self.table)

        self.button = QPushButton("Buscar Tanques")
        self.button.clicked.connect(self.buscar_tanques)
        self.layout.addWidget(self.button)

        self.dados_tanques = []
        self.file_path = "dados_atualizados.xlsx"  # Caminho do arquivo Excel

        # Carregar dados existentes do arquivo Excel, se houver
        if os.path.exists(self.file_path):
            self.load_data_from_excel()

    def buscar_tanques(self):
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
        password_field.send_keys('xpert')

        form = driver.find_element(By.TAG_NAME, 'form')
        form.submit()

        tanques = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.tanque')))

        for tanque in tanques:
            capacidade_element = tanque.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
            litros_element = tanque.find_element(By.CSS_SELECTOR, 'h1.leitura-tanque > b')
            porcentagem_element = tanque.find_element(By.CSS_SELECTOR, 'span.ng-binding')

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
            }

            self.dados_tanques.append(dados)

        driver.quit()

        self.save_data_to_excel()

    def load_data_from_excel(self):
        df = pd.read_excel(self.file_path)
        self.update_table(df)

    def save_data_to_excel(self):
        # Verifica se o arquivo Excel existe
        if os.path.exists(self.file_path):
            # Carrega os dados existentes do Excel para o DataFrame
            df = pd.read_excel(self.file_path)
            # Concatena os novos dados aos dados existentes
            df = pd.concat([df, pd.DataFrame(self.dados_tanques)], ignore_index=True)
        else:
            df = pd.DataFrame(self.dados_tanques)

        # Salva os dados no arquivo Excel
        df.to_excel(self.file_path, index=False)

        # Atualiza a tabela na GUI
        self.update_table(df)

    def update_table(self, df):
        self.table.setRowCount(len(df))

        for row in range(len(df)):
            data_execucao_item = QTableWidgetItem(str(df.loc[row, "Data Execução"]))
            litros_item = QTableWidgetItem(str(df.loc[row, "Litros"]))
            capacidade_item = QTableWidgetItem(str(df.loc[row, "Capacidade"]))
            porcentagem_item = QTableWidgetItem(str(df.loc[row, "Porcentagem"]))

            self.table.setItem(row, 0, data_execucao_item)
            self.table.setItem(row, 1, litros_item)
            self.table.setItem(row, 2, capacidade_item)
            self.table.setItem(row, 3, porcentagem_item)

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    apply_stylesheet(app, theme='dark_teal.xml', invert_secondary=True)

    window.show()
    app.exec()
