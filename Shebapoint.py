import os
import time

from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Get the directory where the script or executable is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "Sheba.xlsx")

# Open the Excel file and read the data into a pandas dataframe
data = read_excel("Sheba.xlsx")
username = "akash.saha@nagad.com.bd"
password = "Black@69"

# Initialize WebDriver
driver = webdriver.Edge()
driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(2)

# Perform login
driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)
try:
    driver.find_element(By.ID, "login_button").click()
except Exception as e:
    print("Login failed:", e)


# Define a function to handle the update process for each row
def process_row(row):
    text_to_match = row['Name']
    business_hours = row['Business Hours']

    try:
        print(f"Processing text: {text_to_match}")

        driver.get('https://sys.mynagad.com:20020/ui/system/#/ntp-management/list')
        time.sleep(2)

        # Click on the element to show 100 items per page
        Click_100_Page = driver.find_element(By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                       "2]/div/div/div/app-list-ntp/section/div["
                                                       "2]/div/div[2]/div["
                                                       "2]/div/app-common-table-advanced/div/div["
                                                       "2]/div/div/button[5]")
        Click_100_Page.click()
        time.sleep(2)

        # Locate the element containing the text to match
        requester_Button_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{text_to_match}')]")))
        requester_Button = requester_Button_element.text

        if requester_Button == text_to_match:
            # Click on the edit button for the found element
            approve_button_locator = (By.XPATH, f"//*[contains(text(), '{text_to_match}')]/following::button["
                                                f"contains(text(), 'Edit')]")
            approve_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(approve_button_locator))
            approve_button.click()
            print("Approved")

            # Enter business hours
            Enter_business_hours = driver.find_element(By.XPATH, '//*[@id="ntpBusinessHours"]')
            Enter_business_hours.clear()
            Enter_business_hours.send_keys(business_hours)

            # Click update button
            Click_update_button = driver.find_element(By.XPATH,
                                                      "/html/body/app-root/app-full-layout/div/div["
                                                      "2]/div/div/div/app-add-ntp/section/div["
                                                      "2]/div/form/div/div/div/div/div["
                                                      "10]/div/div/div/div/div/button[2]")
            Click_update_button.click()

            # Check for success message
            element = driver.find_element(By.ID, value = "toast-container")
            message = element.text
            if "Success!" in message:
                print(f'{text_to_match} is Updated')
                return "Updated"
            else:
                print(f'{text_to_match} is not Updated')
                return "Not Updated"
    except Exception as ff:
        print(f"Could not process {text_to_match}: {ff}")
        return "Error occurred"


# Process each row in the dataframe
data['Status_Message'] = data.apply(process_row, axis = 1)

# Save updated dataframe to Excel
data.to_excel(file_path, index = False, engine = 'openpyxl')

# Close the driver
driver.quit()
