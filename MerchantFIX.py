import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.webdriver import WebDriver

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "Merchant fix.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

driver: WebDriver = webdriver.Edge()
driver.get('https://systest.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()

username = "uatdemo18@gmail.com"
password = "N@gad1234"

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
except:
    print("Login failed......")
    time.sleep(2)

for index, row in data.iterrows():
    driver.get('https://systest.mynagad.com:20020/ui/system/#/merchant-management/list')
    time.sleep(2)

    enter_customer_number = driver.find_element(By.XPATH, '//*[@id="accountNo"]')
    enter_customer_number.clear()
    enter_customer_number.send_keys('0')
    enter_customer_number.send_keys(str(row['Number']))

    search_merchant_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
    search_merchant = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(search_merchant_locator))
    search_merchant.click()

    status_column_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                       "2]/div/div/div/app-merchant-list/section/div/div/div["
                                       "2]/div/div/app-common-table-advanced/table/tbody/tr/td[4]/span/span")
    status_column = WebDriverWait(driver, 3).until(EC.visibility_of_element_located(status_column_locator))

    if status_column.text.strip() == "Active":
        print(f"Status is {status_column.text.strip()}. Proceeding further.")

        click_details = driver.find_element(By.XPATH, "//button[contains(text(), 'Details')]")
        click_details.click()
        time.sleep(2)

        click_edit_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Edit')]")
        click_edit_button.click()
        time.sleep(2)

        checkbox_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                      "2]/div/div/div/app-merchant-reg/section/div[2]/div/form/div/div["
                                      "11]/div/div/div/div[2]/div/label/input")
        checkbox = WebDriverWait(driver, 1).until(EC.presence_of_element_located(checkbox_locator))

        if checkbox.is_selected():
            print("Auto-Settlement is already ticked.")
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "Merchant fix.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[row_number, column_name] = "Auto Settlement is already ticked"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
        else:
            checkbox.click()
            print("Auto-Settlement has been ticked.")
            Enter_period = driver.find_element(by = By.XPATH, value = '//*[@id="settlementPolicy"]')
            Enter_period.send_keys(row['Period'])
            time.sleep(2)
            Select_hour = driver.find_element(by = By.XPATH, value = '//*[@id="hour"]')
            Select_hour.send_keys(row['hour'])
            time.sleep(2)
            select_minute = driver.find_element(by = By.XPATH, value = '//*[@id="minute"]')
            select_minute.send_keys(row['minutes'])
            time.sleep(2)
            select_bank = driver.find_element(by = By.XPATH, value = '//*[@id="autoSettlementBankAccount"]')
            select_bank.send_keys(row['Bank'])
            time.sleep(2)
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "Merchant fix.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[row_number, column_name] = "Auto Settlement is ticked"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
            click_update_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Update')]")
            click_update_button.click()
            time.sleep(2)
            driver.quit()
            driver = webdriver.Edge()
            driver.get('https://systest.mynagad.com:20020/ui/system/#/home')
            driver.maximize_window()

            username = "uatdemo02@gmail.com"
            password = "N@gad1234"

            driver.find_element(By.ID, "username").send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)

            try:
                driver.find_element(By.ID, "login_button").click()
            except:
                print("Login failed......")
                time.sleep(2)

            driver.get('https://systest.mynagad.com:20020/ui/system/#/approval/list')
            time.sleep(2)

            click_merchant_edit = driver.find_element(By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                                "2]/div/div/div/app-approval-task/section/div["
                                                                "2]/div/div/div/div/ngb-accordion/div["
                                                                "10]/div/div/button/div/div["
                                                                "1]/h5/b")
            click_merchant_edit.click()
            time.sleep(2)

            try:
                # Assuming `driver` is your WebDriver instance
                requester_email_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'uatdemo18@gmail.com')]"))
                )
                requester_email = requester_email_element.text

                approve_button_locator = (By.XPATH, "//button[contains(text(), 'Approve')]")
                approve_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(approve_button_locator))

                if requester_email == "uatdemo18@gmail.com":
                    approve_button.click()
                    print("Approved")
                    # If the "Details" button was found, it's a customer number
                    file_path = os.path.join(script_dir, "Merchant fix.xlsx")
                    df = pd.read_excel(file_path)
                    row_number = index  # Assuming you want to start from the first row (index 0)
                    column_name = 'Approval_Status'
                    df.at[row_number, column_name] = "Approved"
                    df.to_excel(file_path, index = False, engine = 'openpyxl')
                    time.sleep(2)
                else:
                    print("Requester email does not match.")
                    time.sleep(2)
                    # If the "Details" button was found, it's a customer number
                    file_path = os.path.join(script_dir, "Merchant fix.xlsx")
                    df = pd.read_excel(file_path)
                    row_number = index  # Assuming you want to start from the first row (index 0)
                    column_name = 'Approval_Status'
                    df.at[row_number, column_name] = "Requester email does not match & could not approve"
                    df.to_excel(file_path, index = False, engine = 'openpyxl')
                    time.sleep(2)
                click_approve_button = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div["
                                                                                  "3]/button[1]")
                click_approve_button.click()
                time.sleep(2)
            except:
                print("Requester email does not match.")
    else:
        print('Not Active Merchant')
