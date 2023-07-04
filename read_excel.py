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
from fastcore.all import parallel

def process_url(row):
    code = str(row.code).strip()

    # Check if 'code' is an empty string, None or 'nan'
    if not code or code.lower() == "nan":
        return None

    # Generate the URL
    url = base_url + code
    print(f"Processing URL: {url}")

    # Load the page
    driver.get(url)

    try:
        # Wait up to 10 seconds until the price table is loaded on the page
        table = WebDriverWait(driver, 0.1).until(
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
                return price

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
            return price

    except Exception as e:
        print(f"Error processing URL: {url}")
        # print(f"Error processing URL: {url}, error: {e}") - Development Only
    return None

# Prompt the user for input and output file names
input_file = input("Enter the name of the xlsx file, including the extension. Example: Test.xlsx: ")
output_file = input("Enter the name of the new xlsx file. If left blank, it will overwrite the input file: ")

# Check if the output file name is empty
if not output_file:
    output_file = input_file
    print(f"Warning: The output file will be overwritten with the same name as the input file: {output_file}")

# Check if the input file exists
if not os.path.isfile(input_file):
    print(f"Error: The input file '{input_file}' does not exist. Make sure the file is in the same directory as the script.")
    exit(1)

# Open the input Excel file
df = pd.read_excel(input_file)

# Configure the Chrome driver
webdriver_service = Service(ChromeDriverManager(cache_valid_range=1).install())
driver = webdriver.Chrome(service=webdriver_service)

# Base URL for product page
base_url = "https://www.digikey.com.br/products/pt?keywords="

# Process each row in the DataFrame
df['price'] = parallel(process_url, df.itertuples(index=False), n_workers=0, progress=True)

# Close the browser
driver.quit()

# Save the DataFrame to the output Excel file
df.to_excel(output_file, index=False, header=True)
