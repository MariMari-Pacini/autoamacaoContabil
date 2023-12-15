import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from time import sleep

# Cria uma nova instância do driver do Google Chrome
service = Service(executable_path='C:\\Users\\marip\\OneDrive\\Área de Trabalho\\chromedriver')
driver = webdriver.Chrome(service=service)

# Carrega o workbook
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Define um dicionário para mapear os nomes das colunas para os IDs dos elementos
element_ids = {
  'product_name': 'product_name',
  'description': 'description',
  'category': 'category',
  'product_code': 'product_code',
  'weight': 'weight',
  'dimensions': 'dimensions',
  'price': 'price',
  'stock': 'stock',
  'expiry_date': 'expiry_date',
  'color': 'color',
  'size': 'size',
  'material': 'material',
  'manufacturer': 'manufacturer',
  'country': 'country',
  'remarks': 'remarks',
  'barcode': 'barcode',
  'warehouse_location': 'warehouse_location'
}

# Define uma função para encontrar um elemento e enviar os valores
def send_keys(element_id, value):
  driver.find_element_by_id(element_id).send_keys(value)

# Define uma função para lidar com a seleção de tamanho
def select_size(tamanho):
  size_select = Select(driver.find_element_by_id('size'))
  if tamanho == 'Pequeno':
      size_select.select_by_visible_text('Small')
  elif tamanho == 'Médio':
      size_select.select_by_visible_text('Medium')
  else:
      size_select.select_by_visible_text('Large')

# Define uma função para lidar com os alertas de confirmação
def handle_confirmation():
  driver.switch_to.alert.accept()
  driver.switch_to.default_content()

# Define uma função para lidar com o processo de registro
def register_product(linha):
  driver.get("https://cadastro-produtos-devaprender.netlify.app/index.html")
  for column, value in zip(element_ids.keys(), linha):
      send_keys(element_ids[column], value)
  driver.find_element_by_xpath('//button[@class="btn btn-primary me-2"]').click()
  sleep(3)
  select_size(linha[10].value)
  driver.find_element_by_xpath('//button[@class="btn btn-primary me-2"]').click()
  sleep(3)
  handle_confirmation()
  handle_confirmation()
  driver.find_element_by_xpath('//button[@class="btn btn btn-primary"]').click()
  sleep(2)

# Itera sobre as linhas na planilha
for linha in sheet_produtos.iter_rows(min_row=2):
  register_product(linha)

# Fecha o navegador
driver.quit()
