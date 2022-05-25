# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import xlsxwriter

workbook = xlsxwriter.Workbook('Lista de Imobiliárias.xlsx')
worksheet = workbook.add_worksheet()
# Abre o Chrome
driver = webdriver.Chrome()
driver.get('https://www.crecisp.gov.br/cidadao/buscarporimobiliaria')

# Espera Resolver o Captcha
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By.XPATH, '/html/body/div/div/section/div[2]'
                                                                       '/div[1]/form/button')))
start = input('Digite 1 para começar:\n')
# Seta o Pages como 1
pages = 1
# Seta as Divs do para clicar
divs = 1
# Seta a Planilha
plan = 2


while pages <= 1069:
    try:
        while divs < 22:
            driver.find_element(By.XPATH, f'/html/body/div/div/section/div[2]/div[{divs}]/form/button').click()
            try:
                object1 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[1]/h3').text
                worksheet.write(f'A{plan}', object1)
                object2 = driver.find_element(By.CLASS_NAME, 'mt-5').text
                worksheet.write(f'B{plan}', object2)
                object3 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[2]').text
                worksheet.write(f'C{plan}', object3)
                object4 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[3]').text
                worksheet.write(f'D{plan}', object4)
                object5 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[4]').text
                worksheet.write(f'E{plan}', object5)
                object6 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[5]').text
                worksheet.write(f'F{plan}', object6)
                object7 = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[6]').text
                worksheet.write(f'G{plan}', object7)
            except: pass
            driver.back()
            divs = divs + 1
            plan = plan + 1
        # Cria Nova aba para manipular as páginas
        driver.get(f'https://www.crecisp.gov.br/cidadao/listadeimobiliarias?page={pages}&IsFinding=True')
        divs = 1  # Reset no Divs
        pages = pages + 1  # Adiciona 1 na próxima página
    except Exception as e:
        print(e)
        print(f'Erro na página{pages}')
        workbook.close()
workbook.close()
