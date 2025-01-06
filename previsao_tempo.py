import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from openpyxl import workbook, load_workbook
from datetime import datetime

# Função para capturar os dados e salvar no Excel
def capturar_dados():
    try:
        # Configuração do Selenium
        navegador = webdriver.Edge()
        navegador.get('https://www.google.com')

        # Busca pela previsão do tempo
        navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('Previsão do tempo')
        navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(Keys.ENTER)

        # Captura dos dados
        temperatura_txt = navegador.find_element(By.XPATH, '//*[@id="wob_tm"]').text
        temperatura = int(temperatura_txt)

        umidade_txt = navegador.find_element(By.XPATH, '//*[@id="wob_hm"]').text
        umidade = int(umidade_txt.replace('%', '').strip())

        # Obter a data e hora atual
        data_hora_atual = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

        # Nome do arquivo Excel
        arquivo_excel = "historico_tempo.xlsx"

        # Trabalhar com a planilha Excel
        try:
            # Tenta abrir a planilha existente
            workbook = load_workbook(arquivo_excel)
            sheet = workbook.active
        except FileNotFoundError:
            # Adiciona cabeçalhos
            sheet.append(["Data/Hora", "Temperatura (°C)", "Umidade (%)"])

        # Adiciona os dados na planilha
        sheet.append([data_hora_atual, temperatura, umidade])

        # Salva o arquivo Excel
        workbook.save(arquivo_excel)
        navegador.quit()

        # Exibe uma mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Dados salvos com sucesso!\nTemperatura: {temperatura}°C\nUmidade: {umidade}%")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        if 'navegador' in locals():
            navegador.quit()

# Criando a interface gráfica com tkinter
def interface():
    janela = tk.Tk()
    janela.title("Coletor de Dados Meteorológicos")
    janela.geometry("400x200")

    # Texto de instrução
    label = tk.Label(janela, text="Clique no botão abaixo para capturar os dados", font=("Arial", 12))
    label.pack(pady=20)

    # Botão para capturar os dados
    botao = tk.Button(janela, text="Capturar Dados", font=("Arial", 14), bg="lightblue", command=capturar_dados)
    botao.pack(pady=20)

    # Iniciar o loop da interface gráfica
    janela.mainloop()

# Executa a interface
interface()


