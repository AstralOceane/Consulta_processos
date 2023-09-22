from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep
import openpyxl

numero_oab = 133864

# Inicializar o WebDriver do Chrome
driver = webdriver.Chrome()
driver.get('https://pje-consulta-publica.tjmg.jus.br/')
sleep(15)

# Digitar o número da OAB
campo_oab = driver.find_element(By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']")
campo_oab.send_keys(numero_oab)

# Selecionar o estado
dropdown_estados = driver.find_element(By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
opcoes_estados = Select(dropdown_estados)
opcoes_estados.select_by_visible_text('SP')

# Clicar em pesquisar
botao_pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
botao_pesquisar.click()
sleep(10)

# Entrar em cada um dos processos
processos = driver.find_elements(By.XPATH, "//b[@class='btn-block']")

# Carregar a planilha ou criar uma nova se não existir
try:
    workbook = openpyxl.load_workbook('dados.xlsx')
except FileNotFoundError:
    workbook = openpyxl.Workbook()

for processo in processos:
    processo.click()
    sleep(10)
    janelas = driver.window_handles
    driver.switch_to.window(janelas[-1])
    driver.set_window_size(1920, 1080)

    # Extrair informações do processo
    numero_processo_element = driver.find_element(By.XPATH, "//div[@class='col-sm-12 ']")
    data_distribuicao_element = driver.find_element(By.XPATH, "//div[@class='value col-sm-12 ']")

    numero_processo = numero_processo_element.text
    data_distribuicao = data_distribuicao_element.text

    movimentacoes = driver.find_elements(By.XPATH, "//div[@id='j_id132:processoEventoPanel_body']//tr[contains(@class,'rich-table-row')]//td//div//div//span")
    lista_movimentacoes = [movimentacao.text for movimentacao in movimentacoes]

    # Criar uma nova planilha para o processo ou acessar uma existente
    try:
        pagina_processo = workbook[numero_processo]
    except KeyError:
        pagina_processo = workbook.create_sheet(numero_processo)

    # Inserir os dados na planilha
    pagina_processo['A1'] = "Número Processo"
    pagina_processo['B1'] = "Data de Distribuição"
    pagina_processo['C1'] = "Movimentações"
    pagina_processo['A2'] = numero_processo
    pagina_processo['B2'] = data_distribuicao

    for index, movimentacao in enumerate(lista_movimentacoes, start=3):
        pagina_processo.cell(row=index, column=3, value=movimentacao)

    # Salvar a planilha
    workbook.save('dados.xlsx')
    driver.close()
    sleep(5)
    driver.switch_to.window(driver.window_handles[0])

# Fechar o WebDriver
driver.quit()
