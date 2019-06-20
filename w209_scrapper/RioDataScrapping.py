## Script para testar as placas de taxi atualmente cadastradas no rio de Janeiro


from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0

import openpyxl

import time

def busca(browser, txtPlaca):
    
    placa = browser.find_element_by_id('lst-ib')
    placa.send_keys(txtPlaca)
    placa.submit()
    resultado = 'zero'
    
    try:
        resultado = browser.find_element_by_id('resultStats').text
    except:
         pass

    return resultado

## Marcando o tempo
start_time = time.time()

## Iniciando o browser para as pesquisas
browser = webdriver.Chrome('/Users/arthurlima/Downloads/chromedriver 2')
browser.get('https://www.google.com.br/')


## Abrindo a planilha de excel para pegar os dados atuais
xl = openpyxl.load_workbook('/Users/arthurlima/Desktop/Python/ed/referral.xlsx')

row = 2

sheet = xl.get_sheet_by_name('Base')

for row in range(2, sheet.max_row + 1):

    val = sheet['A'+str(row)].value

    txtPlaca = val

    
    sheet['B'+str(row)] = busca(browser,txtPlaca)

    
    if row % 500 == 0:

        xl.save('/Users/arthurlima/Desktop/Python/ed/referral.xlsx')
        print row
        end_time = time.time()
        print (end_time - start_time)
    

xl.save('/Users/arthurlima/Desktop/Python/ed/referral.xlsx')

browser.quit()

end_time = time.time()
print (end_time - start_time)
