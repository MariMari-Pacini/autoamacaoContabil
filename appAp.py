import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from time import sleep

# Create a new instance of the Google Chrome driver
service = Service(executable_path='C:\Users\marip\OneDrive\Área de Trabalho\Automacao Contabil')
driver = webdriver.Chrome(service=service)

# Load the workbook
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Iterate over the rows in the sheet
for linha in sheet_produtos.iter_rows(min_row=2):
    # Go to the website
    driver.get("https://cadastro-produtos-devaprender.netlify.app/index.html")

    # Find the elements and send the values
    driver.find_element_by_id('product_name').send_keys(linha[0].value)
    driver.find_element_by_id('description').send_keys(linha[1].value)
    driver.find_element_by_id('category').send_keys(linha[2].value)
    driver.find_element_by_id('product_code').send_keys(linha[3].value)
    driver.find_element_by_id('weight').send_keys(linha[4].value)
    driver.find_element_by_id('dimensions').send_keys(linha[5].value)

    # Click the next button
    driver.find_element_by_xpath('//button[@class="btn btn-primary me-2"]').click()
    sleep(3)

    # Continue sending values
    driver.find_element_by_id('price').send_keys(linha[6].value)
    driver.find_element_by_id('stock').send_keys(linha[7].value)
    driver.find_element_by_id('expiry_date').send_keys(linha[8].value)
    driver.find_element_by_id('color').send_keys(linha[9].value)
    
    # Create a new Select object
    size_select = Select(driver.find_element_by_id('size'))

    # Select an option by visible text
    size_select.select_by_visible_text('Small')

    # Handle the size selection
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        size_select.select_by_visible_text('Small')
    elif tamanho == 'Médio':
        size_select.select_by_visible_text('Medium')
    else:
        size_select.select_by_visible_text('Large')

    # Continue sending values
    driver.find_element_by_id('material').send_keys(linha[11].value)

    # Click the next button
    driver.find_element_by_xpath('//button[@class="btn btn-primary me-2"]').click()
    sleep(3)

    # Continue sending values
    driver.find_element_by_id('manufacturer').send_keys(linha[12].value)
    driver.find_element_by_id('country').send_keys(linha[13].value)
    driver.find_element_by_id('remarks').send_keys(linha[14].value)
    driver.find_element_by_id('barcode').send_keys(linha[15].value)
    driver.find_element_by_id('warehouse_location').send_keys(linha[16].value)

    # Click the finish button
    driver.find_element_by_xpath('//button[@class="btn btn-primary me-2"]').click()
    sleep(3)
    
   # Handle the first confirmation
    driver.switch_to.alert.accept()
    driver.switch_to.default_content()

    # Handle the second confirmation
    driver.switch_to.alert.accept()
    driver.switch_to.default_content()

    # Start the registration again
    driver.find_element_by_xpath('//button[@class="btn btn btn-primary"]').click()
    sleep(2)

# # Close the browser
# driver.quit()
