
# Digikey Price Extractor

This Python script allows you to extract product prices from the DigKey website using an input Bill of Materials (BOM) file in xlsx format. It automates the process of navigating to the product page for each provided code and extracting the corresponding price.

## Requirements

Make sure you have the following packages installed before running the script:

- openpyxl: Library for working with Excel files in xlsx format.
- selenium: Library for browser automation.
- webdriver_manager: Driver manager for Selenium.
- pandas: Library for data manipulation and analysis.
- bs4 (BeautifulSoup): Library for parsing and extracting information from HTML pages.

You can install the packages using pip (Python package manager). Execute the following command to install the required packages:

`pip install openpyxl selenium webdriver_manager pandas bs4
`

## How to Use

1. Run the Python script in an environment that has the mentioned packages installed.
2. When prompted, enter the name of the input file (BOM) in xlsx format, including the extension. For example: Test.xlsx.
3. Next, enter the name of the new xlsx file that will contain the extracted prices. If you leave it blank, the output file will overwrite the input file.
4. Wait while the script processes each code from the Bill of Materials and extracts the corresponding prices.
5. Upon completion, the script will save the output xlsx file with the updated prices.

The table must be like this

| code| price | link |
| :---         |     :---:      |          ---: |
| CODE_OF_PRODUCT  | null     | null   |
|  ...  | ...      | ...     |



## Notes

- Make sure the input file (BOM) is in the same directory as the script.
- The script uses the Chrome WebDriver to automate browsing on the DigKey website. It will automatically download the appropriate driver for the current Chrome version if it's not already installed.
- During execution, the script will display information about the processing of each URL and any encountered errors.
- If a price cannot be extracted for a particular code, the corresponding price field will be left blank in the output file.
- The script is open to suggestions for improvements. Feel free to make suggestions or contribute to the code.
