import os
import time
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# Get the directory where the script or executable is located

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")

    # Open the Excel file and read the data into a pandas dataframe
    data = pd.read_excel(file_path)

    # Start a Chrome WebDriver instance
    driver = webdriver.Chrome()
    driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
    driver.maximize_window()
    time.sleep(2)

    # Login
    username = "sysops.automation@gmail.com"
    password = "Nagad@202404"

    username_field = driver.find_element(By.ID, "username")
    username_field.send_keys(username)

    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys(password)

    login_button = driver.find_element(By.ID, "login_button")
    login_button.click()

    progress_file = "progress.txt"

    # Initialize progress to 0 if progress file does not exist
    if not os.path.exists(progress_file):
        with open(progress_file, "w") as f:
            f.write("0")

    # Load progress from file
    with open(progress_file, "r") as f:
        progress = int(f.read())

    # Wait for login to complete
    WebDriverWait(driver, 10).until(EC.url_contains('home'))

    # Process each row in the dataframe
    for index, row in data.iterrows():
        driver.get('https://sys.mynagad.com:20020/ui/system/#/fee-commission-management/list/biller-merchant')
        time.sleep(2)

        try:
            # Select payee and enter service number
            select_payee = driver.find_element(By.XPATH, '//*[@id="payee"]')
            select_payee.send_keys('UDDOKTA')

            enter_service_number = driver.find_element(By.XPATH, '//*[@id="billerServiceNumber"]')
            enter_service_number.clear()
            import re

            service_number_value = re.sub(r'\.0$', '', str(row[ 'SERVICE_NUMBER' ]))
            enter_service_number.send_keys(service_number_value)

            # Click search
            search_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Search')]")
            search_button.click()
            time.sleep(5)

            # Click details
            details_button = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                        "2]/div/div/div/app-biller-merchant-fee"
                                                                        "-commission-list/app-common-fee-commission-list"
                                                                        "/section/div/div/div["
                                                                        "2]/div/div/app-common-table-advanced/table/tbody"
                                                                        "/tr[1]/td[7]/div/span/button")
            details_button.click()
            time.sleep(5)

        except NoSuchElementException:
            print("No results found. Moving to the next row.")
            continue  # Move to the next iteration of the loop
            print("NO SUCH AG")
            # If the "Details" button was found, it's a customer number
            file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
            df = pd.read_excel(file_path)
            df.at[ index, 'Status_Message' ] = "Not Expired"
            df.to_excel(file_path, index = False, engine = 'openpyxl')
            time.sleep(3)
            driver.refresh()

        try:
            print("Starting APP FEE COMMISSION EXPIRE")
            # Wait for the requester fee to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Uddokta')]")))

            # If requester commission is Uddokta, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire.click()
            time.sleep(2)
            print("Uddokta APP FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Uddokta')]")))

            # If requester commission is Uddokta, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire.click()
            time.sleep(2)
            print("Uddokta APP FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Distributor')]")))

            # If requester commission is Distributor, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_two = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_two.click()
            time.sleep(2)
            print("Distributor APP FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'MD')]")))

            # If requester commission is MD, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_two = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_two.click()
            time.sleep(2)
            print("MD APP COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'TWTL')]")))

            # If requester commission is TWTL, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("TWTL APP COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Bpo')]")))

            # If requester commission is BPO, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("BPO APP COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Advance "
                                                                                      "Commission')]")))

            # If requester commission is Advance Commission, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("ADC APP COM Expired")
            driver.refresh()
            print("APP FEECOM EXPIRE COMPLETED")

            # USSD Configuration #
            # Wait for the requester fee to load
            print("USSD FEECOM EXPIRE STARTED")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Uddokta')]")))

            # If requester commission is Uddokta, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire.click()
            time.sleep(2)
            print("Uddokta USSD FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Uddokta')]")))

            # If requester commission is Uddokta, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire.click()
            time.sleep(2)
            print("Uddokta USSD FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Distributor')]")))

            # If requester commission is Distributor, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_two = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_two.click()
            time.sleep(2)
            print("Distributor USSD FEE COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'MD')]")))

            # If requester commission is MD, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_two = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_two.click()
            time.sleep(2)
            print("MD USSD COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'TWTL')]")))

            # If requester commission is TWTL, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("TWTL USSD COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Bpo')]")))

            # If requester commission is BPO, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("BPO USSD COM Expired")
            driver.refresh()

            # Wait for the requester fee commission to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Advance "
                                                                                      "Commission')]")))

            # If requester commission is Advance Commission, click expire
            expire_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Expire')]")
            expire_button.click()
            time.sleep(5)

            confirm_expire_three = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
            confirm_expire_three.click()
            time.sleep(2)
            print("ADC USSD COM Expired")
            driver.refresh()
            print("USSD FEECOM EXPIRE STARTED")
            element = driver.find_element(By.ID, value = "toast-container")
            message = element.text
            if "Success!" in message:
                print("Expired")
                # If the "Details" button was found, it's a customer number
                file_path = os.path.join(script_dir, "FEE_COM_SUSPENSION.xlsx")
                df = pd.read_excel(file_path)
                row_number = index  # Assuming you want to start from the first row (index 0)
                column_name = 'Status_Message'
                df.at[ row_number, column_name ] = "Expired"
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
            # Update progress in the progress file
            with open(progress_file, "w") as f:
                f.write(str(index + 1))

    except Exception as e:
        print("An error occurred:", e)
        print("Restarting the process...")
        main()
    finally:
        # Save progress
        with open(progress_file, "w") as f:
            f.write(str(index))
        # Save data to excel
        file_path = os.path.join(script_dir, "MRR.xlsx")
        data.to_excel(file_path, index = False, engine = 'openpyxl')
        # Close the WebDriver
    if 'driver' in locals():
        driver.quit()


if __name__ == "__main__":
    main()
