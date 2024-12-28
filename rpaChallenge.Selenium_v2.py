from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin

# Função para lidar com elementos não encontrados
def find_element_safe(driver, by, value, timeout=10):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception:
        print(f"Elemento {value} não encontrado.")
        return None

# Instância do bot
bot = DesktopBot()

# Instância do plugin Excel
bot_excel = BotExcelPlugin()

# Lê os dados do Excel
dados = bot_excel.read('E:\\RPA\\UIPath\\Projetos_estudo\\RPA Challenge\\challenge.xlsx').as_list("Sheet1")[1:]

# Inicia o WebDriver
driver = webdriver.Chrome("E:\\RPA\\RPA_Studio\\DRIVES\\chromedriver-win64\\chromedriver.exe")

# Abre o site
driver.get("https://rpachallenge.com/")

# Aguarda o site carregar
driver.implicitly_wait(10)

# Localiza e clica no botão "Start" para iniciar o desafio
start_button = find_element_safe(driver, By.XPATH, '//button[text()="Start"]')
if start_button:
    start_button.click()

for linha in dados:
    # Variáveis dos campos Excel
    dataFirstname, dataLastname, dataCompanyName, dataRole, dataAdress, dataEmail, dataPhone = linha

    # Preenche os campos do formulário
    find_element_safe(driver, By.XPATH, '//label[text()="First Name"]/following-sibling::input').send_keys(dataFirstname)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelLastName"]').send_keys(dataLastname)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelAddress"]').send_keys(dataAdress)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]').send_keys(dataCompanyName)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelRole"]').send_keys(dataRole)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelPhone"]').send_keys(dataPhone)
    find_element_safe(driver, By.XPATH, '//input[@ng-reflect-name="labelEmail"]').send_keys(dataEmail)

    # Clica no botão "Submit"
    submit_button = find_element_safe(driver, By.XPATH, '//input[@value="Submit"]')
    if submit_button:
        submit_button.click()

# Aguarda o usuário ver o resultado
input("Processo concluído. Pressione Enter para fechar o navegador.")

# Fecha o WebDriver
driver.quit()
