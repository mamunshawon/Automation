import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.webdriver import WebDriver

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

driver: WebDriver = webdriver.Chrome()
driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(2)

username = "mamunur.shawon@nagad.com.bd"
password = "NAgad@112804.."

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
except:
    print("Login failed......")
    time.sleep(2)

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/fee-commission-management/list/biller-merchant')
    time.sleep(2)
    # Find the input field
    enter_service_number = driver.find_element(by = By.XPATH, value = '//*[@id="billerServiceNumber"]')
    # Clear any existing content
    enter_service_number.clear()
    # Get the service number value from the row
    service_number_value = str(row['SERVICE_NUMBER']).strip()
    # Remove trailing zeros
    service_number_value = re.sub(r'0*$', '', service_number_value)
    # Send the desired value
    enter_service_number.send_keys(service_number_value)
    time.sleep(2)
    search_merchant_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
    search_merchant = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(search_merchant_locator))
    search_merchant.click()
    status_column_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                       "2]/div/div/div/app-biller-merchant-fee-commission-list/app-common-fee"
                                       "-commission-list/section/div/div/div["
                                       "2]/div/div/app-common-table-advanced/table/tbody/tr[2]/td[2]/span")
    status_column = WebDriverWait(driver, 3).until(EC.visibility_of_element_located(status_column_locator))
    if status_column.text.strip() == "CU":
        print(f"Status is {status_column.text.strip()}. Proceeding further.")

        click_details = driver.find_element(By.XPATH, "//button[contains(text(), 'Details')]")
        click_details.click()
        time.sleep(2)
        Click_Expire_one = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                      "2]/div/div/div/app-biller-merchant-fee"
                                                                      "-commission-detail/app-common-fee-commission"
                                                                      "-details/section/div/div[1]/div/div["
                                                                      "2]/div/div/form/div/div[2]/div/table/tbody/tr["
                                                                      "1]/td[7]/button")
        Click_Expire_one.click()
        time.sleep(2)
        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)
        Click_Expire_two = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                      "2]/div/div/div/app-biller-merchant-fee"
                                                                      "-commission-detail/app-common-fee-commission"
                                                                      "-details/section/div/div[1]/div/div["
                                                                      "2]/div/div/form/div/div[2]/div/table/tbody/tr["
                                                                      "5]/td[5]/button")
        Click_Expire_two.click()
        time.sleep(2)
        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_three = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                        "2]/div/div/div/app-biller-merchant-fee"
                                                                        "-commission-detail/app-common-fee-commission"
                                                                        "-details/section/div/div[1]/div/div["
                                                                        "2]/div/div/form/div/div["
                                                                        "2]/div/table/tbody/tr[9]/td[4]/button")
        Click_Expire_three.click()
        time.sleep(2)
        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_Four = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-merchant-fee"
                                                                       "-commission-detail/app-common-fee-commission"
                                                                       "-details/section/div/div[1]/div/div["
                                                                       "2]/div/div/form/div/div["
                                                                       "2]/div/table/tbody/tr[13]/td[4]/button")
        Click_Expire_Four.click()
        time.sleep(2)
        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_Five = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-merchant-fee"
                                                                       "-commission-detail/app-common-fee-commission"
                                                                       "-details/section/div/div[1]/div/div["
                                                                       "2]/div/div/form/div/div["
                                                                       "2]/div/table/tbody/tr[17]/td[4]/button")
        Click_Expire_Five.click()
        time.sleep(2)

        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_six = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                      "2]/div/div/div/app-biller-merchant-fee"
                                                                      "-commission-detail/app-common-fee-commission"
                                                                      "-details/section/div/div[1]/div/div["
                                                                      "2]/div/div/form/div/div[2]/div/table/tbody/tr["
                                                                      "21]/td[6]/button")
        Click_Expire_six.click()
        time.sleep(2)

        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_seven = driver.find_element(by = By.XPATH, value = '/html/body/app-root/app-full-layout/div/div['
                                                                        '2]/div/div/div/app-biller-merchant-fee'
                                                                        '-commission-detail/app-common-fee-commission'
                                                                        '-details/section/div/div[1]/div/div['
                                                                        '2]/div/div/form/div/div['
                                                                        '2]/div/table/tbody/tr[25]/td[5]/button')
        Click_Expire_seven.click()
        time.sleep(2)

        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_eight = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                        "2]/div/div/div/app-biller-merchant-fee"
                                                                        "-commission-detail/app-common-fee-commission"
                                                                        "-details/section/div/div[1]/div/div["
                                                                        "2]/div/div/form/div/div["
                                                                        "2]/div/table/tbody/tr[25]/td[5]/button")
        Click_Expire_eight.click()
        time.sleep(2)
        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_nine = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                       "2]/div/div/div/app-biller-merchant-fee"
                                                                       "-commission-detail/app-common-fee-commission"
                                                                       "-details/section/div/div[1]/div/div["
                                                                       "2]/div/div/form/div/div["
                                                                       "2]/div/table/tbody/tr[33]/td[4]/button")
        Click_Expire_nine.click()
        time.sleep(2)

        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)

        Click_Expire_ten = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                      "2]/div/div/div/app-biller-merchant-fee"
                                                                      "-commission-detail/app-common-fee-commission"
                                                                      "-details/section/div/div[1]/div/div["
                                                                      "2]/div/div/form/div/div[2]/div/table/tbody/tr["
                                                                      "37]/td[4]/button")
        Click_Expire_ten.click()
        time.sleep(2)

        click_expire_button = driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")
        click_expire_button.click()
        time.sleep(2)
        element = driver.find_element(By.ID, value = "toast-container")
        message = element.text
        if "Success!" in message:
            print("Expired")
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[row_number, column_name] = "Expired"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
        else:
            print("Not Expired")
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[ row_number, column_name ] = "Not Expired"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
    else:
        print("NO SUCH CU")
        # If the "Details" button was found, it's a customer number
        file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
        df = pd.read_excel(file_path)
        row_number = index  # Assuming you want to start from the first row (index 0)
        column_name = 'Status_Message'
        df.at[row_number, column_name] = "Not Found"
        df.to_excel(file_path, index = False, engine = 'openpyxl')
        time.sleep(3)