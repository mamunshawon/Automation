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
file_path = os.path.join(script_dir, "BookEMI.xlsx")


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
data = read_excel("BookEMI.xlsx")
username = "akash.saha@nagad.com.bd"
password = "Black@69"

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

# Extract row names from the DataFrame
row_names = data.iloc[:, 0].tolist()

# Initialize an empty report string
full_report = ""

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/bill-pay-management/biller-service-detail/161569')
    time.sleep(3)
    click_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                              "2]/div/div/div/app-biller-service-detail/section/div"
                                                              "/div[1]/div/div[2]/div/div/form/div/button[3]")
    click_button.click()
    time.sleep(2)
    Service_Name = driver.find_element(by = By.XPATH, value = "//*[@id='serviceName']")
    Service_Name_Ba = driver.find_element(By.ID, "serviceNameBn")
    Service_Number = driver.find_element(By.ID, "serviceNumber")
    Service_Name.send_keys(row["First Name"])
    time.sleep(2)
    Service_Name_Ba.send_keys(row["Last Name"])
    time.sleep(2)
    Service_Number.send_keys(row["Service Number"])
    time.sleep(4)
    click_merchant_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-service/section/div["
                                                                       "2]/div/div/div/div/form/div/div["
                                                                       "6]/div/div/div/span/i")
    click_merchant_button.click()
    Enter_merchant_number = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                       "-search-merchant/div/div[2]/div["
                                                                       "1]/div/div/form/div[1]/div/div[2]/div/div/input")
    time.sleep(2)
    Enter_merchant_number.send_keys('0')
    Enter_merchant_number.send_keys(row["Merchant Number"])
    time.sleep(5)

    click_search_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[2]/div["
                                                                     "1]/div/div/form/div["
                                                                     "2]/div/button")
    click_search_button.click()
    time.sleep(5)

    click_bullet_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[2]/div["
                                                                     "2]/div/table/tbody/tr/td[4]/input")
    click_bullet_button.click()

    click_select_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[3]/span[2]")
    click_select_button.click()
    time.sleep(5)

    click_edit_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                   "2]/div/div/div/app-biller-service/section/div["
                                                                   "2]/div/div/div/div/form/div/div["
                                                                   "27]/div/table/tbody/tr[1]/td[5]/button")
    click_edit_button.click()
    time.sleep(5)

    # ScrollPage #
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    # ScrollPage #
    driver.execute_script("window.scroll(0.5, 0.5);")
    time.sleep(2)
    Reenter_Merchant_Number = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div"
                                                                         "/div["
                                                                         "2]/div/div/div/app-biller-service/section"
                                                                         "/div[2]/div/div/div/div/form/div/div["
                                                                         "26]/div[2]/div/div/div/div[1]/div["
                                                                         "2]/div/div/div/input")

    Reenter_Merchant_Number.clear()

    driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                               "2]/div/div/div/app-biller-service/section/div["
                                               "2]/div/div/div/div/form/div/div[26]/div[2]/div/div/div/div[1]/div["
                                               "2]/div/div/div/input").send_keys('0')

    driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                               "2]/div/div/div/app-biller-service/section/div["
                                               "2]/div/div/div/div/form/div/div[26]/div[2]/div/div/div/div[1]/div["
                                               "2]/div/div/div/input").send_keys(row["Merchant Number"])
    time.sleep(5)

    click_update_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                     "2]/div/div/div/app-biller-service/section/div["
                                                                     "2]/div/div/div/div/form/div/div[26]/div["
                                                                     "9]/div/button")
    click_update_button.click()
    time.sleep(2)

    click_register_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-service/section"
                                                                       "/div[2]/div/div/div/div/form/div/div["
                                                                       "31]/div/button")
    click_register_button.click()
    time.sleep(2)

    element = driver.find_element(By.ID, value = "toast-container")
    message = element.text
    if "Biller Service added successfully" in message:
        print("Registered")
        file_path = os.path.join(script_dir, "BookEMI.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        print(f'{row["First Name"]} is Registered')
        df.at[row_number, column_name] = "Mentioned Biller is Registered"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
    else:
        print("Not Registered")
        file_path = os.path.join(script_dir, "BookEMI.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        print(f'{row["First Name"]} is Not Registered')
        df.at[row_number, column_name] = "Not Registered"
        df.to_excel(file_path, index = False, engine = 'openpyxl')

    if "Biller Service added successfully" in message:
        status_message = "Registered"
    else:
        status_message = "Not Registered"

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
