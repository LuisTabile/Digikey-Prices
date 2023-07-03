from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from bs4 import BeautifulSoup
import re

# Abre o arquivo xlsx
df = pd.read_excel('Teste.xlsx')

# Configura o driver do Chrome
webdriver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=webdriver_service)

for idx, row in df.iterrows():
    # Gera a URL
    url = f'https://www.digikey.com.br/pt/products/detail/{row["codigo"]}'
    print(f"Processando o URL: {url}")

    # Carrega a página
    driver.get(url)

    try:
        # Espera até 10 segundos até que a tabela de preços seja carregada na página
        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.MuiTable-root.tss-1w92vj0-table.css-u6unfi"))
        )

        # Extrai o código HTML da tabela
        table_html = table.get_attribute('outerHTML')

        # Cria o objeto BeautifulSoup
        soup = BeautifulSoup(table_html, 'html.parser')

        # Encontra todas as linhas da tabela
        rows = soup.find('tbody').find_all('tr')

        # Para cada linha na tabela
        for row in rows:
            # Encontra todas as células na linha
            cells = row.find_all('td')

            # Verifica se a primeira célula contém '1000'
            if len(cells) > 0 and cells[0].text.strip() == '1.000':
                # Remove o símbolo de dólar e espaços do texto
                price_text = cells[1].text.replace('$', '').strip()
                # Substitui a vírgula pelo ponto na string
                price_text = price_text.replace(',', '.')
                # Converte o preço para um número
                price = float(price_text)
                # Atualiza o preço na linha do DataFrame
                df.at[idx, 'preco'] = price
                print(f"Preço atualizado para: {price}")
                break

    except Exception as e:
        print(f"Erro ao processar o URL: {url}, erro: {e}")

# Fecha o navegador
driver.quit()

# Salva o DataFrame alterado de volta para o arquivo xlsx
df.to_excel('Teste.xlsx', index=False)
