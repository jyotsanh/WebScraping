from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.chrome.service import Service
from openpyxl.cell import _writer

# specify the url
url = "https://www.nepalstock.com/floor-sheet"

# specify the path to chromedriver.exe (change it to your own path)
# change path as needed
total_range = int(input("enter the total number of pages to extract : "))

rna = total_range - 4


def extractor(html):
    # Specify the path to your HTML file

    # Parse the HTML using BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')

    whole_table = soup.find(
        'table', class_='table table__lg table-striped table__border table__border--bottom')
    tbody = whole_table.find('tbody')
    tr = tbody.find_all('tr')

    Contract_No = []
    Stock_Symbol = []
    Buyer = []
    Seller = []
    Quantity = []
    Rate = []
    Amount = []

    for i in tr:

        contact = "'" + (i.find_all('td')[1].text)
        Contract_No.append(contact)  # Contract_No
        Stock_Symbol.append(i.find_all('td')[2].text)  # Stock_Symbol
        Buyer.append(i.find_all('td')[3].text)  # Buyer
        Seller.append(i.find_all('td')[4].text)  # Seller
        Quantity.append(i.find_all('td')[5].text)  # Quantity
        Rate.append(i.find_all('td')[6].text)  # Rate
        Amount.append(i.find_all('td')[7].text)  # Amount

    data_dict = {

        'Contract No': Contract_No,
        'Stock Symbol': Stock_Symbol,
        'Buyer': Buyer,
        'Seller': Seller,
        'Quantity': Quantity,
        'Rate': Rate,
        'Amount': Amount
    }
    # df = pd.DataFrame(data_dict)
    # df.to_excel('Stocks.xlsx', index=False)

    file_name = 'Stocks.xlsx'
    # Load existing data from the Excel file, if it exists
    try:
        # Try to read the existing Excel file
        df = pd.read_excel(file_name)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame
        df = pd.DataFrame({
            'Contract No': Contract_No,
            'Stock Symbol': Stock_Symbol,
            'Buyer': Buyer,
            'Seller': Seller,
            'Quantity': Quantity,
            'Rate': Rate,
            'Amount': Amount
        })
    for contacts, symbols, buy, sel, quanti, rate, amt in zip(Contract_No, Stock_Symbol, Buyer, Seller, Quantity, Rate, Amount):
        new_data = pd.DataFrame({
            'Contract No': [contacts],
            'Stock Symbol': [symbols],
            'Buyer': [buy],
            'Seller': [sel],
            'Quantity': [quanti],
            'Rate': [rate],
            'Amount': [amt]
        })
        df = pd.concat([df, new_data], ignore_index=True)
        # df = pd.concat([df, new_data], ignore_index=True)

        # Save the DataFrame to 'report.xlsx' after each iteration
        df.to_excel(file_name, index=False)


path_to_chromedriver = './chromedriver-win64/chromedriver.exe'
service = Service(path_to_chromedriver)
driver = webdriver.Chrome(service=service)


# navigate to the url
driver.get(url)
time.sleep(3)
html_code = driver.page_source
extractor(html_code)
# find the button by its xpath
# replace 'button_xpath' with the actual xpath of the button
button = driver.find_element(
    By.XPATH, "/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[4]/a/span[2]")


# click the button
button.click()
# Retrieve the HTML code
time.sleep(1)
html_code = driver.page_source
extractor(html_code)


button = driver.find_element(
    By.XPATH, "/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[5]/a/span[2]")


# click the button
button.click()
time.sleep(1)
html_code = driver.page_source
extractor(html_code)


button = driver.find_element(
    By.XPATH, "/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[6]/a/span[2]")

# click the button
button.click()
time.sleep(1)
html_code = driver.page_source
extractor(html_code)
if rna <= 0:
    pass
else:
    for i in range(0, rna):

        button = driver.find_element(
            By.XPATH, "/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[7]/a/span[2]")

        # click the button
        button.click()
        time.sleep(1)
        html_code = driver.page_source
        extractor(html_code)

# close the browser
driver.quit()

# /html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[9]/a/span[2]
