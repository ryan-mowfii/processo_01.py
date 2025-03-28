from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from time import sleep
import time
import os
from PIL import Image
from pptx.util import Inches
from pptx import Presentation


                                            # ENTRANDO NA CONCILIADORA


driver = webdriver.Chrome()
driver.maximize_window()

    # --- PASSO 1: Login com múltiplas estratégias ---
driver.get("https://app.conciliadora.com.br/ManagementDashboard")
sleep(5)  # Espera inicial

    # Estratégias para campo de usuário (tentativas em ordem)
driver.find_element(By.ID, 'login').send_keys('contato.audit@up380.com.br')
sleep(2)
driver.find_element(By.ID, 'password').send_keys('Acesso01*')
sleep (2)
driver.find_element(By.ID,'btnLogin').click()
sleep (8)

print('Entrei dentro da conciliadora')

#--------------------------------------------------------------------------------------------------------------------------
                                                       #COMANDO DE LIMPEZA 


limpar_pesquisa_pagamentos = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="dropDownSearch"]/div[1]/div/div[1]'))
    )
limpar_pesquisa_pagamentos.click()
driver.execute_script("arguments[0].focus();", limpar_pesquisa_pagamentos)
sleep(2)
 
    # tecla Backspace 5 vezes usando ActionChains
actions = ActionChains(driver)
for _ in range(5):
        actions.send_keys(Keys.BACK_SPACE)
actions.perform()
 

 #-----------------------------------------------------------------------------------------------------------------------
                                                            # FUNÇÃO DE PESQUISA



def pesquisar_empresa_segura(driver, nome_empresa):
    try:
        print("Iniciando processo de pesquisa seguro...")

        # 3. Preenchimento seguro do campo
        campo = driver.find_element(By.CSS_SELECTOR, '.dx-tag-container input')
        campo.send_keys(nome_empresa)
        time.sleep(1.5)  # Espera generosa para sugestões

        # 4. Seleção sem clicar direto (usando keyboard)
        campo.send_keys(Keys.ARROW_DOWN, Keys.ENTER)
        print(f"✅ '{nome_empresa}' selecionada com sucesso!")
        return True

    except Exception as e:
        print(f"❌ Falha: {str(e)}")
        driver.save_screenshot('erro_pesquisa.png')
        return False

# USO DA FUNÇÃO DE PESQUISA
if pesquisar_empresa_segura(driver, "mercear"):
    print("Processo finalizado!")
else:
    print("Falha - verifique o screenshot")



#-------------------------------------------------------------------------------------------------------------------------
                                                         #COMANDO SELEÇÃO DE DATA
def selecionar_periodo():
    try:
        # 1. Clicar no campo de data para abrir o calendário
        campo_data = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'dateRangePicker'))
        )
        campo_data.click()
        sleep(2)

        # 2. Verificar estrutura do calendário (debug)
        calendario = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'calendar-table'))
        )
        print("Estrutura do calendário encontrada:")
        print(calendario.get_attribute('outerHTML'))

        # 3. Alternativa 1: Seleção por texto visível
        # Data inicial (1)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//td[contains(@class, "available") and text()="1"]'))
        ).click()
        sleep(1)

        # Data final (31)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//td[contains(@class, "available") and text()="31"]'))
        ).click()
        sleep(2)

        # 4. Clicar em aplicar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(@class, "trn") and contains(text(), "Aplicar")]'))
        ).click()
        
        print("Período selecionado com sucesso!")
        return True

    except Exception as e:
        print(f"Erro ao selecionar período: {str(e)}")
        driver.save_screenshot('erro_calendario.png')
        return False

selecionar_periodo()
sleep(5)

#------------------------------------------------------------------------------------------------------------------------
                                                  #COPIAR DO VALOR DE SALDO TOTAL DENTRO DA PAGINA
                                                  

# saldo = driver.find_element(By.XPATH, '//div[@class="dx-datagrid-summary-item dx-datagrid-text-content"]').text
# print(saldo)

def copiar_valor_especifico(driver):
    """
    Copia o valor do elemento especificado na imagem.

    Args:
        driver: Instância do WebDriver do Selenium.

    Returns:
        O valor do elemento como uma string, ou None se o elemento não for encontrado.
    """
    try:
        # Encontra o elemento usando o seletor CSS fornecido
        elemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.dx-datagrid-summary-item.dx-datagrid-text-content[style="text-align: right;"]'))
        )
        # Obtém o texto do elemento
        valor = elemento.text
        return valor
    except Exception as e:
        print(f"Erro ao copiar valor: {e}")
        return None

# Exemplo de uso
# Supondo que você já tenha inicializado o driver e navegado para a página desejada
# driver = webdriver.Chrome()
# driver.get("URL_DA_PAGINA")

valor_copiado = copiar_valor_especifico(driver)

if valor_copiado:
    print(f"Valor copiado: {valor_copiado}")
    # Agora você pode usar 'valor_copiado' para colar em um slide do PowerPoint
else:
    print("Falha ao copiar o valor.")
