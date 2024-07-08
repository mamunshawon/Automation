import os
import time

import pandas as pd
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "CCPortalDeletion.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = read_excel("CCPortalDeletion.xlsx")
username = "akash.saha@nagad.com.bd"
password = "Black@69"

driver: WebDriver = webdriver.Chrome()
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

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/auth-user/list')
    time.sleep(3)

    enter_userID = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                          "2]/div/div/div/app-auth-user-list/section/div/div/div["
                                                          "1]/div/ngb-accordion/div/div[2]/div/div/form/div["
                                                          "1]/div/div[1]/div/input")
    enter_userID.send_keys(row['USER_ID'])
    time.sleep(2)

    click_search_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                                 "2]/div/div/div/app-auth-user-list/section/div/div"
                                                                 "/div[1]/div/ngb-accordion/div/div["
                                                                 "2]/div/div/form/div[2]/div/button")
    click_search_button.click()

    time.sleep(2)

    click_edit_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                               "2]/div/div/div/app-auth-user-list/section/div/div"
                                                               "/div["
                                                               "2]/div/div/app-common-table-advanced/table/tbody/tr"
                                                               "/td[6]/div/span[2]/button")
    click_edit_button.click()
    time.sleep(2)

    select_inactive_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                                    "2]/div/div/div/app-auth-user-edit/section/div["
                                                                    "2]/div/div/div/div/form/div[1]/div["
                                                                    "9]/div/select/option[2]")
    select_inactive_button.click()
    time.sleep(2)

    select_update_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                                  "2]/div/div/div/app-auth-user-edit/section/div["
                                                                  "2]/div/div/div/div/form/div[2]/button[2]")
    select_update_button.click()
    time.sleep(2)
    element = driver.find_element(By.ID, value="toast-container")
    message = element.text
    if "Success!" in message:
        file_path = os.path.join(script_dir, "CCPortalDeletion.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        print(f'{row["USER_ID"]} is Deleted')
        df.at[row_number, column_name] = "User is Deleted"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
    else:
        file_path = os.path.join(script_dir, "CCPortalDeletion.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        print(f'{row["USER_ID"]} is Not Deleted')
        df.at[row_number, column_name] = "User is Not Deleted"
        df.to_excel(file_path, index = False, engine = 'openpyxl')

driver.close()
