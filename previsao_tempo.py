#Biblioteca para trabalhar com automação
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

#Biblioteca para trabalhar com planilhas
from openpyxl import load_workbook

navegador = webdriver.Edge()

#Acessar um site
navegador.get('https://www.google.com')

#Clica nos elementos indicados do site

navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('Previsão do tempo')
navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(Keys.ENTER)

temperatura_str = navegador.find_element(By.XPATH, '//*[@id="wob_tm"]').text  # Use `.text` para capturar o conteúdo visível
temperatura = int(temperatura_str)

umidade_str = navegador.find_element(By.XPATH, '//*[@id="wob_hm"]').text
umidade = int(umidade_str.replace('%', '').strip())


print(f'Temperatura atual: {temperatura} Umidade: {umidade}%' )

#Retorna os nomes das planilhas
#print(arquivo.sheetnames)
