import os
import time

import pandas as pd
import requests
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "PMUK.xlsx")

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
data = read_excel("PMUK.xlsx")
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
except:
    print("login failed.......")
    time.sleep(2)

# Extract row names from the DataFrame
row_names = data.iloc[:, 0].tolist()

# Initialize an empty report string
full_report = ""

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/bill-pay-management/biller-service-detail/195458')
    time.sleep(3)
    click_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                              "2]/div/div/div/app-biller-service-detail/section/div"
                                                              "/div[1]/div/div[2]/div/div/form/div/button[3]")
    click_button.click()
    time.sleep(5)
    Service_Name = driver.find_element(by = By.XPATH, value = "//*[@id='serviceName']")
    Service_Name_Ba = driver.find_element(By.ID, "serviceNameBn")
    Service_Number = driver.find_element(By.ID, "serviceNumber")
    Service_Name.send_keys(row["Branch Name"])
    time.sleep(2)
    Service_Name_Ba.send_keys(row["Branch Name"])
    time.sleep(2)
    Service_Number.send_keys(row["Service ID"])
    time.sleep(2)
    select_date_range = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                   "2]/div/div/div/app-biller-service/section/div["
                                                                   "2]/div/div/div/div/form/div/div[4]/div["
                                                                   "1]/div/div/div")
    select_date_range.click()
    time.sleep(2)

    select_date = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                             "2]/div/div/div/app-biller-service/section/div["
                                                             "2]/div/div/div/div/form/div/div[4]/div["
                                                             "1]/div/ngb-datepicker/div["
                                                             "2]/div/ngb-datepicker-month/div[6]/div[2]/div")
    select_date.click()
    time.sleep(2)

    # enter_time = driver.find_element(by = By.XPATH, value = '//*[@id="availableDateHour"]')
    # enter_time.clear()
    # enter_time.send_keys('0')
    # time.sleep(2)

    click_merchant_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-service/section/div["
                                                                       "2]/div/div/div/div/form/div/div["
                                                                       "6]/div/div/div/span/i")
    click_merchant_button.click()
    time.sleep(2)
    Enter_merchant_number = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                       "-search-merchant/div/div[2]/div["
                                                                       "1]/div/div/form/div[1]/div/div[2]/div/div/input")
    Enter_merchant_number.send_keys('0')
    Enter_merchant_number.send_keys(row["Merchant Number"])
    time.sleep(2)
    click_search_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[2]/div["
                                                                     "1]/div/div/form/div["
                                                                     "2]/div/button")
    click_search_button.click()
    time.sleep(2)

    click_bullet_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[2]/div["
                                                                     "2]/div/table/tbody/tr/td[4]/input")
    click_bullet_button.click()

    click_select_button = driver.find_element(by = By.XPATH, value = "/html/body/ngb-modal-window/div/div/app-common"
                                                                     "-search-merchant/div/div[3]/span[2]")
    click_select_button.click()
    time.sleep(5)
    try:
        requester_Button_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'branch wallet')]"))
        )
        requester_Button = requester_Button_element.text

        if requester_Button == "branch wallet":
            approve_button_locator = (By.XPATH, "//*[contains(text(), "
                                                "'branch wallet')]/following::button[contains("
                                                "text(), 'Edit')]")
            approve_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(approve_button_locator))
            approve_button.click()
            print("Approved")
            time.sleep(2)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            # ScrollPage #
            driver.execute_script("window.scroll(0.5, 0.5);")
            time.sleep(2)
            Reenter_Merchant_Number = driver.find_element(by = By.XPATH, value = '//*[@id="defaultValEn0"]')
            Reenter_Merchant_Number.clear()
            Reenter_Merchant_Number.send_keys('0')
            time.sleep(5)
            Reenter_Merchant_Number.send_keys(row['Merchant Number'])
            time.sleep(5)
            click_update_button = driver.find_element(by = By.XPATH, value = '//*[@id="serviceAttributeInfoUpdate"]')
            click_update_button.click()
            time.sleep(5)

            # Assuming driver is your WebDriver instance

            # Wait for the button to be clickable
            click_register_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                                "/html/body/app-root"
                                                                                                "/app-full-layout/div"
                                                                                                "/div["
                                                                                                "2]/div/div/div/app"
                                                                                                "-biller-service"
                                                                                                "/section/div["
                                                                                                "2]/div/div/div/div"
                                                                                                "/form/div/div["
                                                                                                "31]/div/button")))

            # Scroll into view if necessary
            driver.execute_script("arguments[0].scrollIntoView(true);", click_register_button)

            # Click the button
            click_register_button.click()

        element = driver.find_element(By.ID, value = "toast-container")
        message = element.text
        if "Success!" in message:
            print("Registered")
            file_path = os.path.join(script_dir, "PMUK.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            print(f'{row["Branch Name"]} is Registered')
            df.at[row_number, column_name] = "Mentioned Biller is Registered"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
        else:
            print("Not Registered")
            file_path = os.path.join(script_dir, "PMUK.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            print(f'{row["Branch Name"]} is Not Registered')
            df.at[row_number, column_name] = "Not Registered"
            df.to_excel(file_path, index = False, engine = 'openpyxl')

        if "Biller Service added successfully" in message:
            status_message = "Registered"
        else:
            status_message = "Not Registered"

        # Append the status message to the full report along with the corresponding row name
        row_report = f"Name: {row_names[index]}\nStatus: {status_message}\n\n"
        full_report += row_report
    except:
        print("Not Registered")
# Send the full report to Telegram
telegram_message = f"Full Report:\n{full_report}"
response = send_telegram_message(bot_token, chat_id, telegram_message)
if response.status_code == 200:
    print("Telegram full report sent successfully!")
else:
    print("Failed to send Telegram full report.")
    print(response.text)

driver.close()
