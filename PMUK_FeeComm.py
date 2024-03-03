import os
import time

import pandas as pd
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "PMUK_EMIFeeComm.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = read_excel("PMUK_EMIFeeComm.xlsx")

username = "oliullah.sizan@nagad.com.bd"
password = "N@gad#$1234"

# driver: WebDriver = webdriver.Chrome(executable_path='./driver/chromedriver.exe')

driver: WebDriver = webdriver.Edge()
driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(2)

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
except:
    print("login failed.......")
    time.sleep(2)

import requests

# Define your bot token and chat ID
bot_token = '6510460079:AAEzl9SdC2yKpHCPFfs4f0-Een8k02H3FTc'
chat_id = '-972000340'


def send_telegram_message(bot_token, chat_id, message):
    url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
    data = {
        'chat_id': chat_id,
        'text': message
    }
    response = requests.post(url, data = data)
    return response


# Extract row names from the DataFrame
row_names = data.iloc[:, 0].tolist()

# Initialize an empty report string
full_report = ""

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/fee-commission-management/detail/biller-merchant'
               '/FC0AG0BP04CLO42022149')
    time.sleep(3)

    click_Copy = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                            "2]/div/div/div/app-biller-merchant-fee-commission"
                                                            "-detail/app-common-fee-commission-details/section/div"
                                                            "/div[1]/div/div[2]/div/div/form/div/button")
    click_Copy.click()
    time.sleep(2)

    Enter_Merchant_Number = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-merchant-fee"
                                                                       "-commission"
                                                                       "-create/app-common-fee-commission/section/div["
                                                                       "2]/div/form/div/div/div/div/ngb-accordion/div"
                                                                       "/div[2]/div/div/div[1]/div/div/div/div/span/i")
    Enter_Merchant_Number.click()
    time.sleep(2)
    Enter_Number = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common-search"
                                                              "-merchant/div/div[2]/div[1]/div/div/form/div[1]/div/div["
                                                              "2]/div/div/input")
    time.sleep(2)
    Enter_Number.send_keys('0')
    Enter_Number.send_keys(row["MerchantNumber"])
    time.sleep(2)

    Search_Merchant_Number = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                        "-search-merchant/div/div[2]/div["
                                                                        "1]/div/div/form/div[2]/div/button")
    Search_Merchant_Number.click()
    time.sleep(2)

    select_bullet_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                      "-search-merchant/div/div[2]/div["
                                                                      "2]/div/table/tbody/tr/td[4]/input")
    select_bullet_button.click()
    time.sleep(2)

    select_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common-search"
                                                               "-merchant/div/div[3]/span[2]")
    select_button.click()
    time.sleep(2)

    select_service = driver.find_element(by = By.XPATH, value = '//*[@id="serviceId"]')
    select_service.send_keys(row['Branch'])
    select_service.click()
    time.sleep(2)

    submit_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                               "2]/div/div/div/app-biller-merchant-fee-commission"
                                                               "-create"
                                                               "/app-common-fee-commission/section/div["
                                                               "2]/div/form/div/div/div/div/ngb-accordion/div/div["
                                                               "2]/div/div/div[4]/div/button")
    submit_button.click()
    time.sleep(2)

    select_register_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                        "2]/div/div/div/app-biller-merchant-fee"
                                                                        "-commission-create/app-common-fee-commission"
                                                                        "/section/div[2]/div/form/div["
                                                                        "2]/div/div/app-dynamic-fee-commission/section"
                                                                        "/div[2]/div[3]/div/button")
    select_register_button.click()
    time.sleep(2)
    element = driver.find_element(By.ID, value = "toast-container")
    message = element.text
    if "Fee-Commission registered successfully" in message:
        print("Registered")
        status_message = 'Registered'
        # If the "Details" button was found, it's a customer number
        file_path = os.path.join(script_dir, "PMUK_EMIFeeComm.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'status_message'
        df.at[row_number, column_name] = "Fee-com Registered"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(2)
    else:
        print("Not Registered")
        status_message = 'Not Registered'
        # If the "Details" button was found, it's a customer number
        file_path = os.path.join(script_dir, "PMUK_EMIFeeComm.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'status_message'
        df.at[row_number, column_name] = "Fee-Com Not Registered"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(2)

    # Append the status message to the full report along with the corresponding row name
    row_report = f"Name: {row_names[index]}\nStatus: {status_message}\n\n"
    full_report += row_report

# Send the full report to Telegram
telegram_message = f"FeeCOM:\n{full_report}"
response = send_telegram_message(bot_token, chat_id, telegram_message)
if response.status_code == 200:
    print("Telegram full report sent successfully!")
else:
    print("Failed to send Telegram full report.")
    print(response.text)

driver.close()
