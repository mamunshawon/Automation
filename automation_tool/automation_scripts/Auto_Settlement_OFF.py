import os
import time
import pandas as pd
import logging
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Selenium options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

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


def save_checkpoint(index):
    with open(checkpoint_file_path, "w") as checkpoint_file:
        checkpoint_file.write(str(index))


# Iterate over rows, starting from the last processed index
for index, row in data.iloc[last_processed_index:].iterrows():
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
        WebDriverWait(driver, 10).until(EC.url_matches("https://sys.mynagad.com:20020/ui/system/#/home"))
        driver.maximize_window()
        logger.info(f"Processing row {index} for number {row['Number']}")

        try:
            driver.find_element(By.ID, "username").send_keys("mamunur.shawon@nagad.com.bd")
            time.sleep(2)
            driver.find_element(By.ID, "password").send_keys("NAgad@112804..")
            time.sleep(2)
            driver.find_element(By.ID, "login_button").click()
            time.sleep(5)
        except (NoSuchElementException, WebDriverException) as e:
            logger.error(f"Login failed: {e}")
            continue

        try:
            driver.get('https://sys.mynagad.com:20020/ui/system/#/merchant-management/list')
            time.sleep(5)
            enter_customer_number = driver.find_element(By.XPATH, '//*[@id="username"]')
            enter_customer_number.send_keys(row['Number'])
            time.sleep(2)

            search_merchant = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]"))
            )
            search_merchant.click()
            status_column = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                            "2]/div/div/div/app-merchant-list/section/div/div/div["
                                                            "2]/div/div/app-common-table-advanced/table/tbody/tr/td["
                                                            "4]/span/span"))
            )
        except TimeoutException:
            logger.error("Merchant search timed out.")
            continue

        if status_column.text.strip() == "Active":
            logger.info(f"Status is {status_column.text.strip()}. Proceeding further.")
            try:
                click_details = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Details')]"))
                )
                click_details.click()
                time.sleep(2)

                click_edit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Edit')]"))
                )
                click_edit_button.click()
                time.sleep(2)

                checkbox = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                              "2]/div/div/div/app-merchant-reg/section/div["
                                                              "2]/div/form/div/div[11]/div/div/div/div["
                                                              "2]/div/label/input"))
                )

                if not checkbox.is_selected():
                    checkbox.click()
                    time.sleep(2)
                    Enter_Period = driver.find_element(By.XPATH, '//*[@id="settlementPolicy"]')
                    Enter_Period.send_keys(row['Period'])
                    time.sleep(2)
                    select_hour = driver.find_element(By.XPATH, '//*[@id="hour"]')
                    select_hour.click()
                    select_hour.send_keys(str(row['hour']))
                    time.sleep(2)
                    select_minute = driver.find_element(By.XPATH, '//*[@id="minute"]')
                    select_minute.send_keys(row['minutes'])
                    time.sleep(2)
                    select_bank = driver.find_element(By.XPATH, '//*[@id="autoSettlementBankAccount"]')
                    select_bank.send_keys(row['Bank'])
                    time.sleep(2)
                    Select_Auto_Settlement = driver.find_element(By.XPATH,
                                                                 '//*[@id="autoSettlementBankAccount"]/option[4]')
                    Select_Auto_Settlement.click()
                    time.sleep(2)

                    click_update_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Update')]"))
                    )
                    driver.execute_script("arguments[0].removeAttribute('disabled');", click_update_button)
                    click_update_button.click()
                    logger.info("Merchant update done.")

                    data.at[index, 'Status_Message'] = "Auto Settlement is ticked and time Updated"
                    data.to_excel(file_path, index=False, engine='openpyxl')
                else:
                    logger.info("Auto-Settlement is already ticked.")
                    data.at[index, 'Status_Message'] = "Auto Settlement is already ticked"
                    data.to_excel(file_path, index=False, engine='openpyxl')
            except (TimeoutException, NoSuchElementException) as e:
                logger.error(f"Error while updating merchant: {e}")
                continue
        else:
            logger.info("Not Active Merchant")
            continue

        try:
            driver.quit()
            driver = webdriver.Chrome(options=chrome_options)
            driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
            driver.maximize_window()

            driver.find_element(By.ID, "username").send_keys("sysops.automation@gmail.com")
            driver.find_element(By.ID, "password").send_keys("Nagad@202404")
            driver.find_element(By.ID, "login_button").click()
            time.sleep(3)

            driver.get('https://sys.mynagad.com:20020/ui/system/#/approval/list')
            time.sleep(2)
            click_merchant_edit = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                      "2]/div/div/div/app-approval-task/section/div["
                                                      "2]/div/div/div/div/ngb-accordion/div[6]/div/div/button"))
            )
            click_merchant_edit.click()
            time.sleep(2)
            Click_Page = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                      "2]/div/div/div/app-approval-task/section/div["
                                                      "2]/div/div/div/div/ngb-accordion/div[6]/div["
                                                      "2]/div/app-approval-task-list/app-common-table-advanced/div/div[2]/div/div/button[5]"))
            )
            Click_Page.click()
            time.sleep(2)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            driver.execute_script("window.scroll(0.5, 0.5);")
            time.sleep(2)

            requester_email_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'mamunur.shawon@nagad.com.bd')]"))
            )
            requester_email = requester_email_element.text

            if requester_email == "mamunur.shawon@nagad.com.bd":
                approve_button_locator = (By.XPATH, "//*[contains(text(), 'mamunur.shawon@nagad.com.bd')]/following"
                                                    "::button[contains(text(), 'Approve')]")
                approve_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(approve_button_locator))
                approve_button.click()
                logger.info("Approved")

                Approve_button_Locator = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Yes')]"))
                )
                Approve_button_Locator.click()

                data.at[index, 'Approval_Status'] = "Approved"
                data.to_excel(file_path, index=False, engine='openpyxl')
        except (TimeoutException, NoSuchElementException) as e:
            logger.error(f"Approval process failed: {e}")
            continue

        save_checkpoint(index)

    except Exception as e:
        logger.error(f"An error occurred at index {index}: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
