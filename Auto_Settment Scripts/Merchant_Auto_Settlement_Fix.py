import os
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Configure logging
log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'script.log')
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,  # Set log level to INFO or DEBUG as needed
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "Merchant fix.xlsx")

# Read the last processed index from the Excel file
checkpoint_file_path = os.path.join(script_dir, "checkpoint.txt")
try:
    with open(checkpoint_file_path, "r") as checkpoint_file:
        last_processed_index = int(checkpoint_file.read().strip())
except FileNotFoundError:
    last_processed_index = 0

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

# Iterate over rows, starting from the last processed index
for index, row in data.iloc[last_processed_index:].iterrows():
    try:
        logging.info(f"Processing row {index}: {row}")

        driver = webdriver.Chrome()
        driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
        WebDriverWait(driver, 10).until(EC.url_matches("https://sys.mynagad.com:20020/ui/system/#/home"))
        driver.maximize_window()
        username = "mamunur.shawon@nagad.com.bd"
        password = "NAgad@112804.."
        driver.find_element(By.ID, "username").send_keys(username)
        time.sleep(2)
        driver.find_element(By.ID, "password").send_keys(password)
        time.sleep(2)
        try:
            driver.find_element(By.ID, "login_button").click()
            time.sleep(5)
        except Exception as e:
            logging.warning("Login failed: %s", e)
            time.sleep(2)

        # Your existing code for processing each row goes here...
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
            logging.info(f"Status is {status_column.text.strip()}. Proceeding further.")

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

            if not checkbox.is_selected():
                checkbox.click()
                logging.info("Auto-Settlement is ticked.")
                time.sleep(2)
                # Get the service number value from the row
                Enter_Period = driver.find_element(by=By.XPATH, value='//*[@id="settlementPolicy"]')
                Enter_Period.send_keys(row['Period'])
                time.sleep(2)
                select_hour = driver.find_element(by=By.XPATH, value='//*[@id="hour"]')
                select_hour.click()
                select_hour.send_keys(str(row['hour']))
                time.sleep(2)
                # Optionally, you can wait for some time after selecting an option
                time.sleep(2)
                select_minute = driver.find_element(by=By.XPATH, value='//*[@id="minute"]')
                select_minute.send_keys(row['minutes'])
                time.sleep(2)
                select_bank = driver.find_element(by=By.XPATH, value='//*[@id="autoSettlementBankAccount"]')
                select_bank.send_keys(row['Bank'])
                time.sleep(2)
                Select_Auto_Settlement = driver.find_element(by=By.XPATH, value='//*[@id="autoSettlementBankAccount'
                                                                                '"]/option[4]')
                Select_Auto_Settlement.click()
                time.sleep(2)

                # Click the update button via JavaScript
                try:
                    # Locate the update button
                    click_update_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Update')]")))

                    # Remove the disabled attribute from the button
                    driver.execute_script("arguments[0].removeAttribute('disabled');", click_update_button)

                    # Click the update button
                    click_update_button.click()

                    logging.info("Merchant update done.")

                    # Update status in the Excel file
                    df = pd.read_excel(file_path)
                    df.at[index, 'Status_Message'] = "Auto Settlement is ticked and time Updated"
                    df.to_excel(file_path, index=False, engine='openpyxl')

                except Exception as e:
                    logging.error("Error updating merchant: %s", e)
                driver.close()
                driver = webdriver.Chrome()
                driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
                driver.maximize_window()

                username = "sysops.automation@gmail.com"
                password = "Nagad@202404"

                driver.find_element(By.ID, "username").send_keys(username)
                driver.find_element(By.ID, "password").send_keys(password)

                try:
                    driver.find_element(By.ID, "login_button").click()
                    time.sleep(3)
                except Exception as e:
                    logging.warning("Login failed: %s", e)
                    time.sleep(2)

                driver.get('https://sys.mynagad.com:20020/ui/system/#/approval/list')
                time.sleep(2)

                click_merchant_edit = driver.find_element(By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                                    "2]/div/div/div/app-approval-task/section/div["
                                                                    "2]/div/div/div/div/ngb-accordion/div["
                                                                    "6]/div/div/button")
                click_merchant_edit.click()
                time.sleep(2)
                Click_Page = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                                    "2]/div/div/div/app-approval-task/section/div["
                                                                    "2]/div/div/div/div/ngb-accordion/div[6]/div["
                                                                    "2]/div/app-approval-task-list/app-common-table"
                                                                    "-advanced/div/div[2]/div/div/button[5]")
                Click_Page.click()
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
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
                        Approve_button_Locator = driver.find_element(By.XPATH, "//button[contains(text(), 'Yes')]")
                        approve_Button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(Approve_button_Locator))
                        approve_Button.click()
                        logging.info("Approved")
                        df = pd.read_excel(file_path)
                        row_number = index
                        column_name = 'Approval_Status'
                        df.at[row_number, column_name] = "Approved"
                        df.to_excel(file_path, index=False, engine='openpyxl')
                        driver.close()
                except TimeoutException:
                    logging.error("Timeout occurred while waiting for elements to be located.")
                except Exception as e:
                    logging.error("An error occurred: %s", e)
            else:
                logging.info("Auto-Settlement is already ticked.")
                df = pd.read_excel(file_path)
                row_number = index
                column_name = 'Status_Message'
                df.at[row_number, column_name] = "Auto Settlement is already ticked"
                df.to_excel(file_path, index=False, engine='openpyxl')
                time.sleep(3)
        else:
            logging.info('Not Active Merchant')

        # Update checkpoint after each successful iteration
        with open(checkpoint_file_path, "w") as checkpoint_file:
            checkpoint_file.write(str(index + 1))

    except Exception as e:
        logging.error(f"An error occurred at index {index}: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
