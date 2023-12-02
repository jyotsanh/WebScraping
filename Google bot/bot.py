import tkinter as tk
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
try:
    df = pd.read_csv('Hindi_English_Truncated_Corpus.csv')
    english_sentence = df['english_sentence'].values
except:
    print('error')
# Specify the path to the chromedriver.exe file (download it from https://sites.google.com/chromium.org/driver/)
# Change this to the actual path on your system
chromedriver_path = './chromedriver-win64/chromedriver.exe'

# Create a new Chrome browser instance
driver = webdriver.Chrome(executable_path=chromedriver_path)

# URL to open
url = 'https://translate.google.com/?sl=en&tl=ne&op=translate'

# Open the URL in the browser
driver.get(url)

time.sleep(2)
for i in range(870, len(english_sentence)):

    button = driver.find_element(
        By.XPATH, "/html/body/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[3]/c-wiz[1]/span/span/div/textarea")
    text = english_sentence[i].replace(
        ",", "").replace("'", "").replace('"', '').replace("`", "").replace("+", "").replace("“”", "").replace(".", "").replace("“", "").replace("”", "").replace("-", "")
    button.click()

    # Create a new action chain
    actions = ActionChains(driver)

    # Send CTRL+A to select all text
    actions.key_down(Keys.CONTROL).send_keys(
        'a').key_up(Keys.CONTROL).perform()

    # Send DELETE to remove the selected text
    actions.send_keys(Keys.DELETE).perform()

    button.send_keys(text)
    skip = 10
    time.sleep(skip)

    button = driver.find_element(
        By.XPATH, "/html/body/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[3]/c-wiz[2]/div/div[7]/div/div[8]/div[2]/span[2]/button")
    button.click()

    def get_clipboard():
        root = tk.Tk()
        # this will show no window
        root.withdraw()
        return root.clipboard_get()

    # After clicking the copy button
    translated_text = get_clipboard()

    with open("dataa.txt", 'a', encoding='utf-8') as file:
        sent = text + "," + translated_text.replace(
            ",", "").replace("'", "").replace('"', '').replace("`", "").replace("+", "").replace("“”", "").replace(".", "").replace("“", "").replace("”", "").replace("-", "")
        file.write(sent+"\n")
# Close the browser when done
driver.quit()
# copy =
# write =
