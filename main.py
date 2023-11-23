# Importando libs
from urllib.request import urlopen
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import pandas as pd
import os
import time

# Definindo navegador e url
driver = webdriver.Edge()

# Diretório de download
diretorio_download = "C:\\Users\\marti\\Downloads"

# Diretório do projeto
diretorio_projeto = "C:\\Users\\marti\\OneDrive\\Documentos\\projetos\\python\\testesws\\planilhas"

# Cria um DataFrame vazio para armazenar os dados combinados
dados_combinados = pd.DataFrame()

def exportar_arquivos():
    # Acessa site através do driver
    driver.get('https://www.ssp.sp.gov.br/transparenciassp/consulta.aspx')

    # Clica no botão de furto
    driver.find_element(By.ID, 'cphBody_btnFurtoVeiculo').click()

    # Lista com os ids dos elementos a serem clicados
    id_anos = ["cphBody_lkAno20", "cphBody_lkAno21", "cphBody_lkAno22"]

    # Itera sobre os anos
    for i in range(len(id_anos)):

        # Espera até que o elemento do ano fique visível
        WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.ID, f'{id_anos[i]}')))
        
        # Cria uma instância de ActionChains para lidar com ações de mouse
        actions = ActionChains(driver)
        
        # Encontra o elemento do ano
        element_ano = driver.find_element(By.ID, f'{id_anos[i]}')
        
        # Move o mouse para o elemento do ano
        actions.move_to_element(element_ano).perform()
        
        # Clica no elemento do ano
        element_ano.click()

        # Itera sobre os meses
        for j in range(12):
            # Espera até que o elemento de bloqueio desapareça
            WebDriverWait(driver, 600).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI blockOverlay"))
            )
            # Aguarda alguns segundos antes de clicar no mês
            time.sleep(10)  # 10 segundos

            driver.find_element(By.ID, f'cphBody_listMes{j+1}').click()
            driver.find_element(By.ID, 'cphBody_ExportarBOLink').click()

def mudar_diretorio():
    os.rename(diretorio_download, diretorio_projeto)

def agrupar_arquivos():
    anos = [2020, 2021, 2022]

    # Itera sobre os arquivos no diretório por ano e mês
    for ano in anos:

        # Cria um arquivo para fazer append
        arquivos_unidos = open("Arquivos_unidos.csv","a")

        for mes in range(12):    
            # Abre o arquivo atual
            arquivo = open(f'planilhas/DadosBO_{ano}_{mes+1}(FURTO DE VEÍCULOS).xls')

            # Armazena o arquivo em uma variável
            dados_combinados = arquivo.read()

            # Faz o append do arquivo para o arquivo principal
            arquivos_unidos.write(dados_combinados)

# Fecha o navegador
driver.quit()

# Chamando funções
exportar_arquivos()
mudar_diretorio()
agrupar_arquivos()