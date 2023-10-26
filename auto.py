import time
import pandas as pd


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from config import CHROME_PROFILE_PATH
from datetime import date

#driver
options = webdriver.ChromeOptions()
options.add_argument(CHROME_PROFILE_PATH)
driver = webdriver.Chrome(options=options)
driver.get("https://app.contaazul.com/#/financeiro/contas-a-receber?view=revenue&amp;source=Financeiro%20%3E%20Contas%20a%20Receber&source=Menu%20Principal")

#planilha
data = pd.read_excel('Results.xlsx')
print(data.head)
clientes = data['name']

#inicio da automação
while True:
    try:
        control = driver.find_element(By.XPATH, '//*[@id="statement-list-container"]/table[1]/tbody')
        if (control):
            time.sleep(3)
            break
    except:
        time.sleep(2)

#exibir
driver.find_element(By.XPATH, '//*[@id="type-filter-controller"]/span').click()
time.sleep(.5)

#recebido
driver.find_element(By.XPATH, '//*[@id="typeFilterContainer"]/li[4]/a/span[1]').click()

#aplicar
driver.find_element(By.XPATH, '//*[@id="type-filter"]/ul/li[2]/div/button').click()
time.sleep(3)

#filtrar contas
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/button').click()

#all
#listar a ordem e largar num loop
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/li[1]/a/span[1]').click()
time.sleep(0.3)

#bradesco
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[2]/a/span').click()

#itau
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[3]/a/span').click()

#sap
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[4]/a/span').click()

#sicredi
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[5]/a/span').click()

#nova data
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[16]/a/span').click()

#cartao iugu
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/div/li[35]/a/span').click()

#aplicar
driver.find_element(By.XPATH, '//*[@id="bank-filter"]/ul/li[3]/div/button').click()
time.sleep(3)
 
#ir para cima
driver.find_element(By.CSS_SELECTOR, 'body').send_keys(Keys.CONTROL, Keys.HOME)
time.sleep(1)

#filtrar data
driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[2]/button/span').click()

#hoje
driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[2]/ul/li[1]/a').click()
time.sleep(3)

#ir para dois dias a frente
controle = date.today().weekday()

if controle == 4:
    driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[4]/a[2]/i').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[4]/a[2]/i').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[4]/a[2]/i').click()
    time.sleep(1)
else:
    driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[4]/a[2]/i').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="financeTopFilters"]/div[4]/a[2]/i').click()
    time.sleep(1)

#abrir todos
while True:
    try:
        driver.find_element(By.XPATH, '//*[@id="conteudo"]/div/div[2]/div[2]/div/div[3]/button').click()
        time.sleep(3)
    except:
        break

#skip
def carregando():
    while True:
        loading = driver.find_element(By.XPATH, '//*[@id="loading"]').get_attribute('style')
        if loading == 'display: block;':
            time.sleep(3)
        else:
            break


#pesquisar
def search(cliente):
    pesquisar = driver.find_element(By.XPATH, '//*[@id="textSearch"]')
    pesquisar.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
    try:
        pesquisar.send_keys(cliente)
        pesquisar.send_keys(Keys.ENTER)
        carregando()
    except:
        return
    
    #abrir
    try:
        driver.find_element(By.XPATH, '//*[@id="statement-list-container"]/table[1]/tbody/tr[1]/td[4]/div[1]/span[1]').click()
        time.sleep(3)
    except KeyError:
        print('none')
        return
    
    #trocar conta pra inad
    conta = driver.find_element(By.XPATH, '//*[@id="newIdConta"]')
    conta.click()
    conta.send_keys(Keys.CONTROL, 'A', Keys.DELETE)
    conta.send_keys('01.2 CLIENTE INADIMPLENTE')


for i in range(len(clientes)):
    cliente = clientes[i]
    search(cliente)
    