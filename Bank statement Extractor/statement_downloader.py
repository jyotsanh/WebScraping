import streamlit as st
from datetime import date
from datetime import datetime
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import os

chrome_options = Options()
chrome_options.add_argument('--headless')

chrome_options_muktinath = Options()
chrome_options_muktinath.add_argument("--window-size=0,0") 
 # Set window size as desired (800x600 in this example)
chrome_options_muktinath.add_argument("--window-position=0,1080")  
# Set window position at the lower left corner (0, 1080 in this example)

def prabhu_bank(df, from_date, to_date):
    st.write("Process Initiated")
    cname = df["CNAME"].to_list()

    for i, row in df.iterrows():
        try:
            session = requests.Session()

            usernames = row['username']
            passwords = row['password']
            CNAME = row['CNAME']

            login_url = "https://ibank.prabhubank.com/expressBanking-client-web/api/Customer/auth"
            login_headers = {
                "Content-Type": "application/json;charset=UTF-8",
            }

            login_data = {
                "username": usernames,
                "password": passwords
            }

            login_response = session.post(login_url, headers=login_headers, json=login_data)

            auth_token = login_response.headers["Authorization"]

            ac_number_url = "https://ibank.prabhubank.com/expressBanking-client-web/api/Account/accounts"
            ac_statement_headers = {
                "Content-Type": "application/json;charset=UTF-8",
                "Authorization": auth_token,
                "Sec-Ch-Ua": "Chromium;v=103, .Not/A)Brand;v=99",
                "Sec-Ch-Ua-Mobile": "?0",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.134 Safari/537.36",
                "Sec-Ch-Ua-Platform": "Windows",
                "Origin": "https://ibank.prabhubank.com",
                "Sec-Fetch-Site": "same-origin",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Dest": "empty",
                "Referer": "https://ibank.prabhubank.com",
                "Accept-Encoding": "gzip, deflate",
                "Accept-Language": "en-US,en;q=0.9",
                "Connection": "close"
            }

            ac_number_response = session.get(ac_number_url, headers=ac_statement_headers)
            account_number = ac_number_response.json()
            account_numbers = [account["accountNumber"] for account in account_number]

            folder_path = "Prabhu Bank Statement"
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            for account in account_numbers:
                ac_statement_url = "https://ibank.prabhubank.com/expressBanking-client-web/api/Account/AcStatement"
                ac_statement_data = {
                    "accountNumber": account,
                    "accountingEntry": "ALL",
                    "fromDate": str(from_date),
                    "toDate": str(to_date),
                }

                ac_statement_response = session.post(ac_statement_url, headers=ac_statement_headers, json=ac_statement_data)

                ac_number_response_data = ac_statement_response.json()

                file_name = f"statement_{CNAME}_{account}.xls"
                file_path = folder_path + '/' + file_name

                file_url = "https://ibank.prabhubank.com/expressBanking-client-web/api/Account/downloadAcStatement/EXCEL"

                resp = session.post(file_url, headers=ac_statement_headers, json=ac_statement_data)
                output = open(file_path, "wb")
                output.write(resp.content)
                output.close()

                client_name = cname[i]
                st.text(f"Completed: {client_name}")
                file_name_log = f"log_{default_from_date}.txt"
                file_path_log = os.path.join(folder_path, file_name_log)
                with open(file_path_log, 'a') as file:
                    file.write(f"Completed: {client_name} \n")

        except:
            st.text(f"Failed: {client_name}")
            with open(file_path_log, 'a') as file:
                file.write(f"Not Completed: {client_name} \n")
    # Display completion message
    st.text("Prabhu Bank statements processing completed.")

def muktinath_bank(df, from_date, to_date):
    st.write("Process Initiated")
    cname = df["CNAME"].to_list()
    username = df["username"].to_list()
    pswd = df["password"].to_list()

    for i in range(len(cname)):
        try:
            browser = webdriver.Chrome('driver\chromedriver.exe',options=chrome_options_muktinath)
            browser.get("https://ebanking.muktinathbank.com.np/#/login")

            time.sleep(3)
            credential = browser.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/div/div[2]/div[2]/form/div[1]/div/input")
            credential.send_keys(username[i])

            # browser.find_element(By.NAME, "password").clear() #to clear password placeholder
            credential = browser.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/div/div[2]/div[2]/form/div[2]/div/input")
            credential.send_keys(pswd[i])

            time.sleep(1)
            credential.send_keys(Keys.RETURN)

            time.sleep(5)

            session = requests.Session()

            # Execute JavaScript to retrieve Session Storage data
            session_storage_data = browser.execute_script("return JSON.stringify(window.sessionStorage);")

            # Convert the Session Storage data to a dictionary
            session_storage_dict = eval(session_storage_data)
            # print("hoooo")
            # Get the value of "ngStorage-token" key
            ngstorage_token = session_storage_dict.get("ngStorage-token")
            # print(ngstorage_token)
            # Get all cookies
            cookies = browser.get_cookies()
            
            browser.quit()

            # Print the "name" and "value" attributes of each cookie
            for cookie in cookies:
                name = cookie['name']
                value = cookie['value']
                # print(f"{name}={value}")
                
                excel_download_url = "https://ebanking.muktinathbank.com.np/expressBanking-client-web/api/Account/downloadAcStatement/EXCEL"

                time.sleep(2)
                # print("uuuuuhikhiwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww")

                account_number_url = "https://ebanking.muktinathbank.com.np/expressBanking-client-web/api/Account/accounts"
                
                excel_download_headers = {

                    "Accept": "application/json, text/plain, */*",
                    "Authorization": ngstorage_token,
                    "Cookie": f"{name}={value}",
                    "Accept": "application/json, text/plain, */*",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Connection": "keep-alive",
                    "Content-Length": "110",
                    "Content-Type": "application/json;charset=UTF-8",
                    "Host": "ebanking.muktinathbank.com.np",
                    "Origin": "https://ebanking.muktinathbank.com.np",
                    "Referer": "https://ebanking.muktinathbank.com.np/",
                    "Sec-Ch-Ua": "\"Not.A/Brand\";v=\"8\", \"Chromium\";v=\"114\", \"Brave\";v=\"114\"",
                    "Sec-Ch-Ua-Mobile": "?0",
                    "Sec-Ch-Ua-Platform": "\"Windows\"",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-origin",
                    "Sec-Gpc": "1",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"

                        }
                # print("lplplpllp[l]")

                
                
                #raj Start

                # Account statement request
                ac_statement_url = "https://ebanking.muktinathbank.com.np/expressBanking-client-web/api/Account/accounts"
                ac_statement_headers = {
                        "Content-Type": "application/json;charset=UTF-8",
                        "Authorization": ngstorage_token,
                        "Sec-Ch-Ua": "Chromium;v=103, .Not/A)Brand;v=99",
                        "Sec-Ch-Ua-Mobile": "?0",
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.134 Safari/537.36",
                        "Sec-Ch-Ua-Platform": "Windows",
                        "Origin": "https://ebanking.muktinathbank.com.np",
                        "Sec-Fetch-Site": "same-origin",
                        "Sec-Fetch-Mode": "cors",
                        "Sec-Fetch-Dest": "empty",
                        "Referer": "https://ebanking.muktinathbank.com.np/",
                        "Accept-Encoding": "gzip, deflate",
                        "Accept-Language": "en-US,en;q=0.9",
                        "Connection": "close"   
                                        }
                
                ac_statement_response = session.get(ac_statement_url, headers=ac_statement_headers)

                # Extract the balance amount
                account = ac_statement_response.json()[0]['accountNumber']

                #raj end


                excel_download_payload = {
                    "accountNumber" : str(account),
                    "accountingEntry" :  "ALL",
                    "fromDate" :   str(from_date),
                    "toDate" : str(to_date)
                }
                # print("-----------------------")
                time.sleep(2)
                download_excel = requests.post(url=excel_download_url, headers=excel_download_headers, json=excel_download_payload)
                # print(download_excel.content)
            folder_path = "Muktinath Statements"
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            file_name = f"statement_{cname[i]}.xls"
            file_path = folder_path + '/' + file_name

            # Check if the request was successful
            if download_excel.status_code == 200:
                # Save the downloaded file
                with open(file_path, "wb") as f:
                    f.write(download_excel.content)
                # print("Excel file downloaded successfully.")
                st.text(f"Completed: {cname[i]}")
                # st.write(f"{cname[i]} completed")
                client_name = cname[i]
                file_name_log = f"log_{default_from_date}.txt"
                file_path_log = os.path.join(folder_path, file_name_log)
                with open(file_path_log, 'a') as file:
                    file.write(f"Completed: {client_name} \n")
            else:
                print("Failed to download the Excel file. Status code:", download_excel.content)
        except:
            st.text(f"Failed: {client_name}")
            with open(file_path_log, 'a') as file:
                file.write(f"Not Completed: {client_name} \n")

    # Display completion message
    st.text("Muktinath Bank statements download completed.")

def mahalaxmi_bank(df, from_date, to_date):
    st.write("Process Initiated")
    if from_date >= to_date:
        st.error("'From' date should be before 'To' date.")
        return
    cname = df["CNAME"].to_list()
    username = df["username"].to_list()
    pswd = df["password"].to_list()
    # Set up ChromeDriver options
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')

    # Specify the path to the ChromeDriver executable
    driver_path = 'driver/chromedriver.exe'

    # Create a new ChromeDriver instance
    driver = webdriver.Chrome(driver_path, options=chrome_options)

    for i in range(len(cname)):
        try:
            # Open the desired URL
            driver.get("https://ibank.mahalaxmibank.com/login.asp")

            # Wait for the pop-up to appear (adjust the sleep duration as needed)
            time.sleep(4)

            # Simulate double-clicking anywhere on the screen
            actions = ActionChains(driver)
            actions.move_by_offset(1, 1).click().click().perform()
            actions.move_by_offset(1, 1).click().click().perform()

            # Input box
            usrname = username[i]  # Replace with your actual username
            password = pswd[i]  # Replace with your actual password

            time.sleep(2)
            username_field = driver.find_element(By.NAME, 'CustCode')
            password_field = driver.find_element(By.NAME, 'password')

            # Enter the username and password
            username_field.send_keys(usrname)
            password_field.send_keys(password)

            time.sleep(2)
            # Click the login button
            driver.find_element(By.NAME, 'image1').click()

            driver.find_element(By.CSS_SELECTOR, 'a[href="InetCustInfoCaller.asp?ForceLogin=T"]').click()
            time.sleep(2)
            # Get all cookies and print their values
            cookies = driver.get_cookies()
            cookie_string = ""
            for index, cookie in enumerate(cookies):
                cookie_string += f"{cookie['name']}={cookie['value']}"
                if index != len(cookies) - 1:
                    cookie_string += "; "

            time.sleep(3)

            url = 'https://ibank.mahalaxmibank.com/pages/InetAcStmnt.asp'

            headers = {
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept-Language': 'en-US,en;q=0.9',
                'Cache-Control': 'max-age=0',
                'Connection': 'keep-alive',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': cookie_string,
                'Host': 'ibank.mahalaxmibank.com',
                'Origin': 'https://ibank.mahalaxmibank.com',
                'Referer': 'https://ibank.mahalaxmibank.com/pages/InetaskdateORG.asp?value=statement',
                'Sec-Ch-Ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Brave";v="114"',
                'Sec-Ch-Ua-Mobile': '?0',
                'Sec-Ch-Ua-Platform': '"Windows"',
                'Sec-Fetch-Dest': 'frame',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'same-origin',
                'Sec-Fetch-User': '?1',
                'Sec-Gpc': '1',
                'Upgrade-Insecure-Requests': '1',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
            }

            payload = {
                'month': str(from_date.month),
                'date': str(from_date.day),
                'year': str(from_date.year),
                'hidden': '',
                'sendDate': f'{from_date.month}/{from_date.day}/{from_date.year}',
                'month1': str(to_date.month),
                'date1': str(to_date.day),
                'year1': str(to_date.year),
                'hidden1': '',
                'sendDate1': f'{to_date.month}/{to_date.day}/{to_date.year}'
            }

            response = requests.post(url, headers=headers, data=payload)

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                table = soup.find('table')
                if table is not None:
                    rows = table.find_all('tr')
                    data = []
                    for row in rows[1:]:
                        cells = row.find_all('td')
                        data.append([cell.get_text(strip=True) for cell in cells])

                    headers = ['TranDate', 'Description', 'ValueDate', 'Debit', 'Credit', 'Balance']
                    df = pd.DataFrame(data, columns=headers)
                    # df.to_excel('table.xlsx', index=False)
                    # print("Table exported to table.xlsx")
                    folder_path = "Mahalaxmi Bank Statement"
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    client_name = cname[i]
                    file_name = f"statement_{client_name}.xlsx"
                    file_name_log = f"log_{default_from_date}.txt"
                    file_path = os.path.join(folder_path, file_name)
                    file_path_log = os.path.join(folder_path, file_name_log)
                    df.to_excel(file_path, index=False)
                    st.text(f"Completed: {client_name}")
                    with open(file_path_log, 'a') as file:
                        file.write(f"Completed: {client_name} \n")
                else:
                    print("Table not found in the response")
            else:
                print('Request failed with status code:', response.text)
        
        except:
            client_name = cname[i]
            st.text(f"Failed: {client_name}")
            with open(file_path_log, 'a') as file:
                file.write(f"Not Completed: {client_name} \n")
    st.text("Mahalaxmi Bank statements download completed.")

if __name__ == "__main__":
    st.title("Statement Downloader")

    # Option selection
    options = ['Prabhu Bank', 'Muktinath Bank','Mahalaxmi Bank']
    option = st.selectbox("Select Bank", options)
    st.write('You selected:', option)

    # Date inputs
    current_datetime = datetime.now()
    default_from_date = date.today()
    default_to_date = date.today()
    from_date = st.date_input("From", default_from_date)
    to_date = st.date_input("To", default_to_date)

    # File upload
    xlsx_file_upload = st.file_uploader("Upload XLSX", type=["xlsx"])

    # Check if file is uploaded and show submit button
    if xlsx_file_upload is not None:
        df = pd.read_excel(xlsx_file_upload)

        if option == 'Prabhu Bank':
            submit_button = st.button("Submit")
            if submit_button:
                prabhu_bank(df, from_date, to_date)
        elif option == 'Muktinath Bank':
            submit_button = st.button("Submit")
            if submit_button:
                muktinath_bank(df, from_date, to_date)

        elif option == 'Mahalaxmi Bank':
            submit_button = st.button("Submit")
            if submit_button:
                mahalaxmi_bank(df, from_date, to_date)
    else:
        st.write("Please upload the file")
