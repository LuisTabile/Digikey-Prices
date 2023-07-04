import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from bs4 import BeautifulSoup

# Prompt the user for input and output file names
input_file = input("Digite o nome do arquivo xlsx, junto com a extensão. Exemplo: Teste.xlsx: ")
output_file = input("Digite o nome do novo arquivo xlsx. Se deixar em branco, irá sobrescrever o arquivo de entrada: ")

# Check if the output file name is empty
if not output_file:
    output_file = input_file
    print(f"Aviso: O arquivo de saída será sobrescrito com o mesmo nome do arquivo de entrada: {output_file}")

# Check if the input file exists
if not os.path.isfile(input_file):
    print(f"Erro: O arquivo de entrada '{input_file}' não existe. Certifique-se de que o arquivo esteja na mesma pasta que o script.")
    exit(1)

# Open the input Excel file
df = pd.read_excel(input_file)

# Configure the Chrome driver
webdriver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=webdriver_service)

# Base URL for product page
base_url = "https://www.digikey.com.br/products/pt?keywords="

for idx, row in df.iterrows():
    code = str(row["code"]).strip()

    # Check if 'code' is an empty string, None or 'nan'
    if not code or code.lower() == "nan":
        continue

    # Generate the URL
    url = base_url + code
    print(f"Processing URL: {url}")

    # Update the link in the DataFrame
    df.at[idx, 'link'] = url

    # Load the page
    driver.get(url)

    try:
        # Wait up to 10 seconds until the price table is loaded on the page
        table = WebDriverWait(driver, 0.5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.MuiTable-root.tss-1w92vj0-table.css-u6unfi"))
        )

        # Extract the HTML code of the table
        table_html = table.get_attribute('outerHTML')

        # Create the BeautifulSoup object
        soup = BeautifulSoup(table_html, 'html.parser')

        # Find all rows in the table
        rows = soup.find('tbody').find_all('tr')

        # Search for the cell containing '1.000'
        price_found = False
        for row in rows:
            # Find all cells in the row
            cells = row.find_all('td')

            # Check if the first cell contains '1.000'
            if len(cells) > 0 and cells[0].text.strip() == '100.000':
                # Remove the dollar symbol and spaces from the text
                price_text = cells[1].text.replace('$', '').strip()
                # Replace the comma with a dot in the string
                price_text = price_text.replace(',', '.')
                # Convert the price to a number
                price = float(price_text)
                # Update the price in the DataFrame row
                df.at[idx, 'preco'] = price
                print(f"Updated price: {price}")
                price_found = True
                break

        if not price_found:
            # Find the last cell in the middle column
            cells = soup.select("table.MuiTable-root.tss-1w92vj0-table.css-u6unfi td:nth-child(2)")
            if cells:
                # Get the last cell in the middle column
                last_cell = cells[-1]
                # Remove the dollar symbol and spaces from the text
                price_text = last_cell.text.replace('$', '').strip()
                # Replace the comma with a dot in the string
                price_text = price_text.replace(',', '.')
                # Convert the price to a number
                price = float(price_text)
                # Update the price in the DataFrame row
                df.at[idx, 'preco'] = price
                print(f"Updated price: {price}")

    except Exception as e:
        print(f"Error processing URL: {url}, error: {e}")

# Close the browser
driver.quit()

# Create a new workbook
workbook = Workbook()
# Create a new sheet in the workbook
sheet = workbook.active
# Save the modified DataFrame to the sheet
for row in dataframe_to_rows(df, index=False, header=True):
    sheet.append(row)

# Save the workbook to the output Excel file
workbook.save(output_file)
