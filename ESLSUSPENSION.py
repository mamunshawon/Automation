import os
import re
import time
import pandas as pd
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

# GET THE DIRECTORY
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "SUSPENSION.xlsx")

# Open the data in EXCEL FILE
data = read_excel("SUSPENSION.xlsx")
username = "mamunur.shawon@nagad.com.bd"
password = "NAgad@112804.."

driver: WebDriver = webdriver.Chrome()
driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(2)

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
    time.sleep(2)
    print("Login Successful")
except:
    print("Login Failed")
    time.sleep(3)

row_names = data.iloc[:, 0].tolist()

for index, row in data.iterrows():
    driver.get("https://sys.mynagad.com:20020/ui/system/#/bill-pay-management/biller-service-list")
    time.sleep(2)

    # Find the input field
    enter_service_number = driver.find_element(by = By.XPATH, value = '//*[@id="serviceNumber"]')

    # Clear any existing content
    enter_service_number.clear()

    # Get the service number value from the row
    service_number_value = str(row['SERVICE_NUMBER']).strip()

    # Remove trailing zeros
    service_number_value = re.sub(r'0*$', '', service_number_value)

    # Send the desired value
    enter_service_number.send_keys(service_number_value)

    time.sleep(2)
    search_biller_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
    search_biller = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(search_biller_locator))
    search_biller.click()
    status_column_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                       "2]/div/div/div/app-biller-service-list/section/div/div/div["
                                       "2]/div/div/app-common-table-advanced/table/tbody/tr/td[5]/span/span")
    status_column = WebDriverWait(driver, 3).until(EC.visibility_of_element_located(status_column_locator))
    if status_column.text.strip() == "Active":
        print(f"Status is {status_column.text.strip()}. Proceeding further.")
        click_details = driver.find_element(By.XPATH, "//button[contains(text(), 'Details')]")
        click_details.click()
        time.sleep(5)
        click_suspend_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Suspend')]")
        click_suspend_button.click()
        time.sleep(2)
        input_reason = driver.find_element(by = By.XPATH, value = '/html/body/div[3]/div/div[2]/textarea')
        input_reason.send_keys('Request from KAM')
        time.sleep(2)
        click_confirm_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Confirm')]")
        click_confirm_button.click()
        time.sleep(2)
        file_path = os.path.join(script_dir, "SUSPENSION.xlsx")
        df = pd.read_excel(file_path)
        row_number = index
        column_name = 'Status_Message'
        df.at[row_number, column_name] = "Biller_Suspended"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        driver.refresh()
    else:
        print("Biller is already Suspended")
        file_path = os.path.join(script_dir, "SUSPENSION.xlsx")
        df = pd.read_excel(file_path)
        row_number = index
        column_name = 'Status_Message'
        df.at[row_number, column_name] = "Biller_already_Suspended"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(2)
        driver.refresh()
