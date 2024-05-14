import os
import time
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def login(driver, username, password):
    driver.get('https://systest.mynagad.com:20020/ui/system/#/home')
    driver.maximize_window()

    username_field = driver.find_element(By.ID, "username")
    username_field.send_keys(username)

    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys(password)

    login_button = driver.find_element(By.ID, "login_button")
    login_button.click()

    WebDriverWait(driver, 10).until(EC.url_contains('home'))


def process_row(driver, row):
    driver.get('https://systest.mynagad.com:20020/ui/system/#/fee-commission-management/list/biller-merchant')
    time.sleep(5)

    try:
        # Select payee and enter service number
        select_payee = driver.find_element(By.XPATH, '//*[@id="payee"]')
        select_payee.send_keys('UDDOKTA')

        enter_service_number = driver.find_element(By.XPATH, '//*[@id="billerServiceNumber"]')
        enter_service_number.clear()
        service_number_value = str(row['SERVICE_NUMBER']).rstrip('.0')
        enter_service_number.send_keys(service_number_value)

        # Click search
        search_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]")
        search_button.click()

        # Click details
        details_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Details')]")))
        details_button.click()

    except NoSuchElementException:
        print("No fee commission found for service number:", service_number_value)
        return False
    except Exception as e:
        print("An error occurred while processing row:", str(e))
        return False

    return True


def expire_fee_com(driver):
    # Expire fee commissions
    fee_types = ['Uddokta', 'Distributor', 'MD', 'TWTL', 'Bpo', 'Advance Commission']
    for fee_type in fee_types:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{fee_type}')]")))
            expire_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Expire')]")))
            expire_button.click()
            confirm_expire = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[3]/button[1]")))
            confirm_expire.click()
            print(f"{fee_type} APP COM Expired")
        except Exception as e:
            print(f"Error occurred while expiring {fee_type} APP COM:", str(e))
            # Optionally, you can implement retry logic here
            # For now, just continue to the next fee_type
            continue


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
    data = pd.read_excel(file_path)

    driver = webdriver.Edge()

    username = "uatdemo18@gmail.com"
    password = "N@gad1234"
    login(driver, username, password)

    for index, row in data.iterrows():
        if process_row(driver, row):
            expire_fee_com(driver)
            # Update status message in Excel
            data.at[index, 'Status_Message'] = "Expired"
            data.to_excel(file_path, index=False, engine='openpyxl')

    driver.quit()


if __name__ == "__main__":
    main()

