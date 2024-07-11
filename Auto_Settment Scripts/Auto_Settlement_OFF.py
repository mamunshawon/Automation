import os
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# Configure Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Constants
LOGIN_URL = 'https://sys.mynagad.com:20020/ui/system/#/home'
MERCHANT_MANAGEMENT_URL = 'https://sys.mynagad.com:20020/ui/system/#/merchant-management/list'
APPROVAL_URL = 'https://sys.mynagad.com:20020/ui/system/#/approval/list'
LOGIN_TIMEOUT = 20
ELEMENT_TIMEOUT = 10
RETRY_ATTEMPTS = 3
RETRY_DELAY = 5
HEADLESS = True

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "AutoSettlementOFF.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

# Selenium options for headless mode
chrome_options = Options()
if HEADLESS:
    chrome_options.add_argument("--headless")  # Run in headless mode


def login(driver, username, password):
    try:
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.ID, "login_button").click()
        WebDriverWait(driver, LOGIN_TIMEOUT).until(EC.url_contains('system'))
        time.sleep(2)  # Allow time for any redirects or UI changes after login
        logger.info("Login successful")
        return True
    except TimeoutException:
        logger.error("Login timeout")
    except Exception as e:
        logger.error("Login failed", exc_info=True)
    return False


def navigate_to(driver, url):
    try:
        driver.get(url)
        WebDriverWait(driver, ELEMENT_TIMEOUT).until(EC.url_contains(url.split('/')[-1]))
        time.sleep(2)  # Allow time for the page to load
        logger.info(f"Successfully navigated to {url}")
        return True
    except TimeoutException:
        logger.error(f"Timeout occurred while navigating to {url}")
    except Exception as e:
        logger.error(f"Navigation to {url} failed", exc_info=True)
    return False


def search_merchant(driver, customer_number):
    try:
        enter_customer_number = driver.find_element(By.XPATH, '//*[@id="username"]')
        enter_customer_number.send_keys(customer_number)
        time.sleep(2)
        search_merchant_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
        search_merchant = WebDriverWait(driver, ELEMENT_TIMEOUT).until(EC.element_to_be_clickable(search_merchant_locator))
        search_merchant.click()
        logger.info("Merchant search successful")
        return True
    except TimeoutException:
        logger.error("Timeout occurred while searching for merchant")
    except Exception as e:
        logger.error("Search merchant failed", exc_info=True)
    return False


def untick_auto_settlement(driver, index, file_path):
    try:
        status_column_locator = (By.XPATH,
                                 "/html/body/app-root/app-full-layout/div/div["
                                 "2]/div/div/div/app-merchant-list/section/div/div/div["
                                 "2]/div/div/app-common-table-advanced/table/tbody/tr/td[4]/span/span")
        status_column = WebDriverWait(driver, ELEMENT_TIMEOUT).until(EC.visibility_of_element_located(status_column_locator))

        if status_column.text.strip() == "Active":
            logger.info(f"Status is {status_column.text.strip()}. Proceeding further.")
            click_details = driver.find_element(By.XPATH, "//button[contains(text(), 'Details')]")
            click_details.click()
            time.sleep(2)

            click_edit_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Edit')]")
            click_edit_button.click()
            time.sleep(2)

            checkbox_locator = (By.XPATH,
                                "/html/body/app-root/app-full-layout/div/div["
                                "2]/div/div/div/app-merchant-reg/section/div[2]/div/form/div/div[11]/div/div/div/div["
                                "2]/div/label/input")
            checkbox = WebDriverWait(driver, 1).until(EC.presence_of_element_located(checkbox_locator))

            if checkbox.is_selected():
                checkbox.click()
                try:
                    click_update_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Update')]")))
                    driver.execute_script("arguments[0].removeAttribute('disabled')", click_update_button)
                    driver.execute_script("arguments[0].click();", click_update_button)
                    logger.info('Auto-Settlement is unticked')
                except Exception as e:
                    logger.error("Error updating merchant", exc_info=True)

                df = pd.read_excel(file_path)
                df.at[index, 'Status_Message'] = "Auto Settlement is unticked"
                df.to_excel(file_path, index=False, engine='openpyxl')
                time.sleep(3)
                return True
            else:
                logger.info("Auto-Settlement is already unticked.")
                df = pd.read_excel(file_path)
                df.at[index, 'Status_Message'] = "Auto Settlement is already unticked"
                df.to_excel(file_path, index=False, engine='openpyxl')
                time.sleep(3)
        else:
            logger.info('Not Active Merchant')
            return False
    except TimeoutException:
        logger.error("Timeout occurred while waiting for elements to be located.")
    except Exception as e:
        logger.error("Failed to untick auto settlement", exc_info=True)
    return False


def approve_changes(driver, index, file_path):
    try:
        navigate_to(driver, APPROVAL_URL)
        time.sleep(2)

        click_merchant_edit = driver.find_element(By.XPATH,
                                                  "/html/body/app-root/app-full-layout/div/div["
                                                  "2]/div/div/div/app-approval-task/section/div["
                                                  "2]/div/div/div/div/ngb-accordion/div[6]/div/div/button")
        click_merchant_edit.click()
        time.sleep(5)

        Click_Page = driver.find_element(By.XPATH,
                                         "/html/body/app-root/app-full-layout/div/div["
                                         "2]/div/div/div/app-approval-task/section/div["
                                         "2]/div/div/div/div/ngb-accordion/div[6]/div["
                                         "2]/div/app-approval-task-list/app-common-table-advanced/div/div["
                                         "2]/div/div/button[5]")
        Click_Page.click()
        time.sleep(2)

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.execute_script("window.scroll(0.5, 0.5);")
        time.sleep(2)

        requester_email_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'mamunur.shawon@nagad.com.bd')]")))
        requester_email = requester_email_element.text

        if requester_email == "mamunur.shawon@nagad.com.bd":
            approve_button_locator = (By.XPATH,
                                      "//*[contains(text(), 'mamunur.shawon@nagad.com.bd')]/following::button["
                                      "contains(text(), 'Approve')]")
            approve_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(approve_button_locator))
            approve_button.click()
            logger.info("Approved")

            Approve_button_Locator = driver.find_element(By.XPATH, "//button[contains(text(), 'Yes')]")
            approve_Button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(Approve_button_Locator))
            approve_Button.click()

            df = pd.read_excel(file_path)
            df.at[index, 'Approval_Status'] = "Approved"
            df.to_excel(file_path, index=False, engine='openpyxl')
            time.sleep(2)
            return True
    except TimeoutException:
        logger.error("Timeout occurred while waiting for elements to be located.")
    except Exception as e:
        logger.error("An error occurred", exc_info=True)
    return False


def perform_operation(driver, username, password, customer_number, index):
    try:
        if not driver:
            driver = webdriver.Chrome(options=chrome_options)

        login(driver, username, password)
        navigate_to(driver, MERCHANT_MANAGEMENT_URL)
        search_merchant(driver, customer_number)

        if untick_auto_settlement(driver, index, file_path):
            driver.close()
            driver = None  # Reset driver to None for new login

            driver = webdriver.Chrome(options=chrome_options)
            login(driver, "sysops.automation@gmail.com", "Nagad@202404")
            approve_changes(driver, index, file_path)
    except Exception as e:
        logger.error("Exception occurred during operation", exc_info=True)
    finally:
        if driver:
            driver.quit()


# Perform operations for each row in the data
for index, row in data.iterrows():
    driver = None
    perform_operation(driver, "mamunur.shawon@nagad.com.bd", "NAgad@112804..", row['Number'], index)
