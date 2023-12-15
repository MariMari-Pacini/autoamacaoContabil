# Projeto de Automação de Registro de Produtos com Selenium
Este projeto é uma versão aprimorada do projeto original do canal Dev Aprender no YouTube (App.py), que usava pyautogui e pyperclip. O objetivo deste projeto é automatizar o processo de registro de produtos em um site, usando a biblioteca Selenium para Python (Application.py).

# Pré-requisitos
Antes de começar, certifique-se de ter o seguinte instalado em seu sistema:

- Python 3.x
- Selenium
- openpyxl
- Google Chrome
- ChromeDriver (compatível com a versão do seu Google Chrome)

# Como usar
Baixe a planilha produtos_ficticios.xlsx e coloque-a na mesma pasta do script Python.
Baixe o driver do Google Chrome (ChromeDriver) que é compatível com a versão do seu Google Chrome e coloque-o na mesma pasta do script Python.
Execute o script Python. O script irá ler a planilha e registrar cada produto no site.

# Estrutura do Código
## O script Python é dividido em várias funções para lidar com diferentes partes do processo de registro de produtos:

> send_keys(element_id, value): Encontra um elemento pelo seu ID e envia valores para ele.
> select_size(tamanho): Lida com a seleção do tamanho do produto.
> handle_confirmation(): Lida com os alertas de confirmação.
> register_product(linha): Lida com todo o processo de registro de um produto.

# Contribuições
Contribuições são bem-vindas! Se você encontrar algum bug ou tiver alguma sugestão de melhoria, sinta-se à vontade para abrir uma issue ou enviar um pull request.
