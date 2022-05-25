# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import xlsxwriter
workbook = xlsxwriter.Workbook('Lista de Corretores.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Nome do Corretor')
worksheet.write('B1', 'Creci')
worksheet.write('C1', 'Telefone')

driver = webdriver.Chrome()
# Abre o Chrome
driver.get('https://www.crecisp.gov.br/cidadao/buscaporcorretores')

# Espera Resolver o Captcha
WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, '/html/body/div/div/section/div[2]'
                                                                      '/div[1]/form/button')))
cont_situacao = 1
pages = 2312
plan = 2

run = input('Digite Qualquer coisa para iniciar')

while pages <= 5000:
    try:
        while cont_situacao < 22:
            # Verifica Situação do Corretor
            situacao = driver.find_element(By.XPATH, f'/html/body/div/div/section/div[2]/div[{cont_situacao}]/div[3]/span').text
            if situacao == 'Ativo':
                # Clica no botão "Ver Detalhes"
                driver.find_element(By.XPATH, f'/html/body/div/div/section/div[2]/div[{cont_situacao}]/form/button').click()
                nome_corretor = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[2]/h3').text
                creci_corretor = driver.find_element(By.XPATH,'/html/body/div/div/section/div[1]/div[2]/div/div[1]/span').text
                print(nome_corretor)
                print(f'Creci: {creci_corretor}')
                worksheet.write(f'A{plan}', nome_corretor)
                worksheet.write(f'B{plan}', creci_corretor)
                try:
                    telefone_corretor = driver.find_element(By.XPATH,'/html/body/div/div/section/div[1]/div[2]/div/div[6]/div/span').text
                    print(telefone_corretor)
                    worksheet.write(f'C{plan}', telefone_corretor)
                except Exception:
                    try:
                        telefone_corretor = driver.find_element(By. XPATH, '/html/body/div/div/section/div[1]/div[2]/div'
                                                                            '/div[5]/div/span').text
                        print(telefone_corretor)
                        worksheet.write(f'C{plan}', telefone_corretor)
                    except Exception:
                        try:
                            telefone_corretor = driver.find_element(By.XPATH, '/html/body/div/div/section/div[1]/div[2]/div'
                                                                            '/div[7]/div/span').text
                            print(telefone_corretor)
                            worksheet.write(f'C{plan}', telefone_corretor)
                        except Exception:
                            try:
                                telefone_corretor = driver.find_element(By.XPATH,
                                                                        '/html/body/div/div/section/div[1]/div[2]/div'
                                                                        '/div[4]/div/span').text
                                print(telefone_corretor)
                                worksheet.write(f'C{plan}', telefone_corretor)
                            except Exception:
                                print('Sem Informações de Telefone')
                                worksheet.write(f'C{plan}', 'Sem telefone')
                plan = plan + 1
                driver.back()
                cont_situacao = cont_situacao + 1
            else:
                # Pula para o próximo corretor
                cont_situacao = cont_situacao + 1
        driver.get(f'https://www.crecisp.gov.br/cidadao/listadecorretores?page={pages}')
        cont_situacao = 1
        pages = pages + 1
    except Exception:
        sleep(1000)
        pages = pages + 1
        print(f'Página Travada!: {pages}')
        pages = pages + 1
        driver.get(f'https://www.crecisp.gov.br/cidadao/listadecorretores?page={pages}')
        try:
            workbook.close()
        except Exception as e:
            print(e)
        cont_situacao = 1
workbook.close()
driver.close()
print('Fim do Programa')
