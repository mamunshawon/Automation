import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.webdriver import WebDriver

def initialize_driver():
    return webdriver.Edge()

def login(driver, username, password):
    driver: WebDriver = webdriver.Edge()
    driver.get('https://systest.mynagad.com:20020/ui/system/#/home')
    driver.maximize_window()
    username = "uatdemo18@gmail.com"
    password = "N@gad1234"
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    try:
        driver.find_element(By.ID, "login_button").click()
        time.sleep(2)
    except:
        print("Login failed......")
        time.sleep(2)

def search_merchant(driver, customer_number):
    # Implementation for searching merchant
    enter_customer_number = driver.find_element(By.XPATH, '//*[@id="accountNo"]')
    enter_customer_number.clear()
    enter_customer_number.send_keys('0')
    enter_customer_number.send_keys(str(row['Number']))

    search_merchant_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
    search_merchant = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(search_merchant_locator))
    search_merchant.click()

def update_merchant(driver, row):
    # Implementation for updating merchant details

def approve_merchant(driver, requester_email):
    # Implementation for approving merchant

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "Merchant fix.xlsx")
    data = pd.read_excel(file_path)

    for index, row in data.iterrows():
        driver = initialize_driver()

        try:
            login(driver, "uatdemo18@gmail.com", "N@gad1234")
            search_merchant(driver, str(row['Number']))

            # Other actions (update, approve) can be called here using functions

        except Exception as e:
            print(f"Error: {str(e)}")

        finally:
            driver.quit()

if __name__ == "__main__":
    main()
