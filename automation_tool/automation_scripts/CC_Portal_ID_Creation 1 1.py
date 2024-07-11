import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Constants
USERNAME = "akash.saha@nagad.com.bd"
PASSWORD = "Black@69"
URL_LOGIN = 'https://sys.mynagad.com:20020/ui/system/#/home'
URL_ADD_USER = 'https://sys.mynagad.com:20020/ui/system/#/auth-user/add'
EXCEL_FILE = "CCPortalCreation.xlsx"
STATUS_MESSAGE_COLUMN = 'Status_Message'
TIMEOUT = 10

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, EXCEL_FILE)

# Open the Excel file and read the data into a pandas dataframe
data = pd.read_excel(file_path)

# Setup WebDriver
driver: WebDriver = webdriver.Chrome()
driver.maximize_window()


# Login Function
def login(driver, username, password):
    driver.get(URL_LOGIN)
    WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.ID, "login_button").click()


# Add User Function
def add_user(driver, row):
    driver.get(URL_ADD_USER)
    WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.XPATH, "//*[@id='userId']"))).send_keys(
        row['mailAddress'])

    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[2]/div/ss-multiselect-dropdown/div/button").click()
    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[2]/div/ss-multiselect-dropdown/div/div/a[3]/span/span["
                        "2]").click()

    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[3]/div/ss-multiselect-dropdown/div/button").click()
    role_input = driver.find_element(By.XPATH, "//*[@id='roleInfo']/div/div/div/input")
    role_input.send_keys('CO')
    WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='roleInfo']/div/div/a[2]/span/span[2]"))).click()

    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[4]/div/input").send_keys(
        row["Name"])
    contact_number = driver.find_element(By.XPATH,
                                         "/html/body/app-root/app-full-layout/div/div["
                                         "2]/div/div/div/app-auth-user-create/section/div["
                                         "2]/div/div/div/div/form/div[1]/div[5]/div/input")
    contact_number.send_keys('0')
    contact_number.send_keys(row["Contact No"])
    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[6]/div/input").send_keys(
        row["Organization Name"])
    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[7]/div/input").send_keys(
        row["Designation"])
    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[1]/div[8]/div/input").send_keys(
        row["Department"])
    driver.find_element(By.XPATH,
                        "/html/body/app-root/app-full-layout/div/div[2]/div/div/div/app-auth-user-create/section/div["
                        "2]/div/div/div/div/form/div[2]/div/button").click()

    # Check for success message
    try:
        WebDriverWait(driver, TIMEOUT).until(EC.presence_of_element_located((By.ID, "toast-container")))
        message = driver.find_element(By.ID, "toast-container").text
        if "Success!" in message:
            return True
        else:
            return False
    except:
        return False


# Main Script
try:
    login(driver, USERNAME, PASSWORD)
    for index, row in data.iterrows():
        success = add_user(driver, row)
        if success:
            print(f'{row["Name"]} is Registered')
            data.at[index, STATUS_MESSAGE_COLUMN] = "User is Added"
        else:
            print(f'{row["Name"]} is not Registered')
            data.at[index, STATUS_MESSAGE_COLUMN] = "User is Not Added"
        driver.refresh()
finally:
    driver.quit()

# Save the updated DataFrame back to the Excel file
data.to_excel(file_path, index=False, engine='openpyxl')
