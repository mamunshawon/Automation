import os
import time
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.webdriver import WebDriver

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "Merchant fixx.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

for index, row in data.iterrows():
    driver: WebDriver = webdriver.Edge()
    driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
    driver.maximize_window()
    time.sleep(2)

    username = "mamunur.shawon@nagad.com.bd"
    password = "NAgad@112804.."

    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)

    try:
        driver.find_element(By.ID, "login_button").click()
        time.sleep(5)
    except:
        print("Login failed......")
        time.sleep(5)
    driver.get('https://sys.mynagad.com:20020/ui/system/#/merchant-management/list')
    time.sleep(5)

    enter_customer_number = driver.find_element(By.XPATH, '//*[@id="username"]')
    enter_customer_number.send_keys(row['Number'])
    time.sleep(2)

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

        # Assuming you already have instantiated the driver and located the checkbox element
        if checkbox.is_selected():
            checkbox.click()
            try:
                # Locate the update button
                click_update_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Update')]"))
                )
                # Execute JavaScript to enable the button and then click it
                driver.execute_script("arguments[0].removeAttribute('disabled')", click_update_button)
                driver.execute_script("arguments[0].click();", click_update_button)
                print('Auto-Settlement is unticked')
            except Exception as e:
                print("Error:", e)
            file_path = os.path.join(script_dir, "Merchant fixx.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[row_number, column_name] = "Auto Settlement is unticked"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
            driver.close()
            driver = webdriver.Edge()
            driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
            driver.maximize_window()

            username = "uatdemo04@gmail.com"
            password = "4hE08J64"

            driver.find_element(By.ID, "username").send_keys(username)
            driver.find_element(By.ID, "password").send_keys(password)

            try:
                driver.find_element(By.ID, "login_button").click()
                time.sleep(5)
            except:
                print("Login failed......")
                time.sleep(5)

            driver.get('https://sys.mynagad.com:20020/ui/system/#/approval/list')
            time.sleep(2)

            click_merchant_edit = driver.find_element(By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                                "2]/div/div/div/app-approval-task/section/div["
                                                                "2]/div/div/div/div/ngb-accordion/div["
                                                                "6]/div/div/button")
            click_merchant_edit.click()
            time.sleep(5)
            Click_Page = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                    "2]/div/div/div/app-approval-task/section/div["
                                                                    "2]/div/div/div/div/ngb-accordion/div[6]/div["
                                                                    "2]/div/app-approval-task-list/app-common-table"
                                                                    "-advanced/div/div[2]/div/div/button[5]")
            Click_Page.click()
            time.sleep(2)
            # ScrollPage #
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            # ScrollPage #
            driver.execute_script("window.scroll(0.5, 0.5);")
            time.sleep(2)
            try:
                requester_email_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'mamunur.shawon@nagad.com.bd')]"))
                )
                requester_email = requester_email_element.text

                if requester_email == "mamunur.shawon@nagad.com.bd":
                    approve_button_locator = (By.XPATH, "//*[contains(text(), "
                                                        "'mamunur.shawon@nagad.com.bd')]/following::button[contains("
                                                        "text(), 'Approve')]")
                    approve_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(approve_button_locator))
                    approve_button.click()
                    print("Approved")

                    Approve_button_Locator = driver.find_element(By.XPATH, "//button[contains(text(), 'Yes')]")
                    approve_Button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(Approve_button_Locator))
                    approve_Button.click()

                    file_path = os.path.join(script_dir, "Merchant fixx.xlsx")
                    df = pd.read_excel(file_path)
                    row_number = index  # Assuming you want to start from the first row (index 0)
                    column_name = 'Approval_Status'
                    df.at[row_number, column_name] = "Approved"
                    df.to_excel(file_path, index = False, engine = 'openpyxl')
                    time.sleep(2)
                    driver.close()
            except TimeoutException:
                print("Timeout occurred while waiting for elements to be located.")
            except Exception as e:
                print("An error occurred:", e)
        else:
            print("Auto-Settlement could not be unticked/already ticked.")
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "Merchant fixx.xlsx")
            df = pd.read_excel(file_path)
            row_number = index  # Assuming you want to start from the first row (index 0)
            column_name = 'Status_Message'
            df.at[row_number, column_name] = "Auto Settlement could not be unticked"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
    else:
        print('Not Active Merchant')
