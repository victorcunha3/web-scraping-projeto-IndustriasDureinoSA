import csv
import re
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog, QTableWidgetItem, QTableWidget
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Busca de Tanques")
        self.setMinimumSize(400, 300)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.label = QLabel("Dados dos Tanques:")
        self.layout.addWidget(self.label)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Litros", "Capacidade", "Porcentagem"])
        self.layout.addWidget(self.table)

        self.button = QPushButton("Buscar Tanques")
        self.button.clicked.connect(self.buscar_tanques)
        self.layout.addWidget(self.button)

        self.download_button = QPushButton("Exportar(EXCEL)")
        self.download_button.clicked.connect(self.download_data)
        self.download_button.setEnabled(False)
        self.layout.addWidget(self.download_button)

        self.dados_tanques = []

    def buscar_tanques(self):
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
        password_field.send_keys('xpert')

        form = driver.find_element(By.TAG_NAME, 'form')
        form.submit()

        tanques = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.tanque')))

        self.dados_tanques = []

        for tanque in tanques:
            capacidade_element = tanque.find_element(By.XPATH, './/h4[contains(text(), "Capacidade")]/b')
            litros_element = tanque.find_element(By.CSS_SELECTOR, 'h1.leitura-tanque > b')
            porcentagem_element = tanque.find_element(By.CSS_SELECTOR, 'span.ng-binding')

            capacidade_text = capacidade_element.text
            capacidade_match = re.search(r'[\d.]+', capacidade_text)
            capacidade = capacidade_match.group() if capacidade_match else ''

            litros_text = litros_element.text
            litros_match = re.search(r'[\d.]+', litros_text)
            litros = litros_match.group() if litros_match else ''
            porcentagem = porcentagem_element.text

            dados = {
                "Litros": litros,
                "Capacidade": capacidade,
                "Porcentagem": porcentagem,
            }

            self.dados_tanques.append(dados)

        driver.quit()

        self.table.setRowCount(len(self.dados_tanques))

        for row, dados in enumerate(self.dados_tanques):
            litros_item = QTableWidgetItem(dados["Litros"])
            capacidade_item = QTableWidgetItem(dados["Capacidade"])
            porcentagem_item = QTableWidgetItem(dados["Porcentagem"])

            self.table.setItem(row, 0, litros_item)
            self.table.setItem(row, 1, capacidade_item)
            self.table.setItem(row, 2, porcentagem_item)

        self.download_button.setEnabled(True)

    def download_data(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getSaveFileName(self, "Salvar arquivo", "", "Arquivo XLSX (*.xlsx)")

        if file_path:
            df = pd.DataFrame(self.dados_tanques)
            df.to_excel(file_path, index=False)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
