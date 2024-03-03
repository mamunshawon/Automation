import os
import time

import pandas as pd
import requests
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "PMUK_Suspend.xlsx")

# Define your bot token and chat ID
bot_token = '6510460079:AAEzl9SdC2yKpHCPFfs4f0-Een8k02H3FTc'
chat_id = '-972000340'

driver: WebDriver = webdriver.Chrome()


def send_telegram_message(bot_token, chat_id, message):
    url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
    data = {
        'chat_id': chat_id,
        'text': message
    }
    response = requests.post(url, data = data)
    return response


# Open the Excel file and read the data into a pandas dataframe
data = read_excel("PMUK_Suspend.xlsx")
username = "mamunur.shawon@nagad.com.bd"
password = "NAgad@112804.."

driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(2)

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
    time.sleep(5)
    print("Login Successful")
except:
    print("login failed.......")
    time.sleep(2)

# Extract row names from the DataFrame
row_names = data.iloc[:, 0].tolist()

# Initialize an empty report string
full_report = ""

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/bill-pay-management/biller-service-list')
    time.sleep(5)
    enter_service_number = driver.find_element(By.XPATH, '//*[@id="serviceNumber"]')
    enter_service_number.clear()
    service_number = str(row['Biller_Num']).rstrip('0')  # Remove trailing zeros
    enter_service_number.send_keys(service_number)
    time.sleep(2)
    click_Search_button = driver.find_element(by = By.XPATH, value = '//*[@id="btnSearchId"]')
    click_Search_button.click()
    time.sleep(1)
    click_details_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Details')]")
    click_details_button.click()
    time.sleep(2)
    click_suspension_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Suspend')]")
    click_suspension_button.click()
    time.sleep(2)
    input_Reason = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[2]/textarea")
    input_Reason.send_keys('Request from Business')
    time.sleep(2)
    click_confirm = driver.find_element(by = By.XPATH, value = "//button[contains(text(),'Confirm')]")
    click_confirm.click()
    time.sleep(2)
    element = driver.find_element(By.ID, value = "toast-container")
    message = element.text
    if "Success!" in message:
        print("Suspended")
        # If the "Details" button was found, it's a customer number
        file_path = os.path.join(script_dir, "PMUK_Suspend.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'status_message'
        df.at[row_number, column_name] = "Suspended"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(2)
        driver.refresh()
    else:
        print("Not Suspended")
        file_path = os.path.join(script_dir, "PMUK_Suspend.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        df.at[row_number, column_name] = "Not Suspended"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(2)
        driver.refresh()

    if "Success!" in message:
        status_message = "Suspended"
    else:
        status_message = "Not Suspended"

    # Append the status message to the full report along with the corresponding row name
    row_report = f"Name: {row_names[index]}\nStatus: {status_message}\n\n"
    full_report += row_report

# Send the full report to Telegram
telegram_message = f"Full Report:\n{full_report}"
response = send_telegram_message(bot_token, chat_id, telegram_message)
if response.status_code == 200:
    print("Telegram full report sent successfully!")
else:
    print("Failed to send Telegram full report.")
    print(response.text)
driver.close()