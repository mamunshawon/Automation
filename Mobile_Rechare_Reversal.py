import os
import time

import pandas as pd
from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


def main():
    # GET THE DIRECTORY
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "MRR.xlsx")

    # Open the data in EXCEL FILE
    data = read_excel("MRR.xlsx")
    username = "mamunur.shawon@nagad.com.bd"
    password = "NAgad@112804.."

    progress_file = "progress.txt"

    # Initialize progress to 0 if progress file does not exist
    if not os.path.exists(progress_file):
        with open(progress_file, "w") as f:
            f.write("0")

    # Load progress from file
    with open(progress_file, "r") as f:
        progress = int(f.read())

    try:
        driver: WebDriver = webdriver.Chrome()
        driver.get('https://sys.mynagad.com:20020/ui/system/#/home')
        driver.maximize_window()
        time.sleep(2)

        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)

        driver.find_element(By.ID, "login_button").click()
        time.sleep(2)
        print("Login Successful")

        row_names = data.iloc[:, 0].tolist()

        for index, row in data.iloc[ progress: ].iterrows():
            driver.get("https://sys.mynagad.com:20020/ui/system/#/transaction-management/reversal/list")
            time.sleep(2)
            Enter_transaction_ID = driver.find_element(by = By.XPATH, value = '//*[@id="transactionId"]')
            Enter_transaction_ID.send_keys(row[ 'Transaction ID' ])
            time.sleep(2)
            Enter_Date = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                    "2]/div/div/div/app-transaction-reversal-list"
                                                                    "/section/div"
                                                                    "/div/div/div/ngb-accordion/div/div["
                                                                    "2]/div/div/form/div["
                                                                    "1]/div/div[2]/div/div/div/div")
            Enter_Date.click()
            time.sleep(1)
            select_date = driver.find_element(by = By.XPATH, value = "/html/body/app-root/app-full-layout/div/div["
                                                                     "2]/div/div/div/app-transaction-reversal-list"
                                                                     "/section"
                                                                     "/div/div/div/div/ngb-accordion/div/div["
                                                                     "2]/div/div/form/div[1]/div/div["
                                                                     "2]/div/div/ngb-datepicker/div["
                                                                     "2]/div/ngb-datepicker-month/div[5]/div[2]/div")
            select_date.click()
            time.sleep(1)
            search_biller_locator = (By.XPATH, "//button[contains(text(), 'Search')]")
            search_biller = WebDriverWait(driver, 1).until(EC.element_to_be_clickable(search_biller_locator))
            search_biller.click()
            time.sleep(10)
            Click_partial_reversal_button = driver.find_element(by = By.XPATH,
                                                                value = "/html/body/app-root/app-full-layout/div"
                                                                        "/div["
                                                                        "2]/div/div/div/app-transaction-reversal"
                                                                        "-list/section/div/div/div["
                                                                        "2]/div/div/div/app-common-table"
                                                                        "-advanced"
                                                                        "/table/tbody/tr/td[6]/div/span["
                                                                        "2]/button")
            time.sleep(2)

            if Click_partial_reversal_button.is_displayed():
                Click_partial_reversal_button.click()
                Enter_Reason = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[2]/textarea")
                time.sleep(2)
                Enter_Reason.send_keys("Reverse")
                Click_yes_button = driver.find_element(by = By.XPATH, value = "/html/body/div[3]/div/div[3]/button[1]")
                Click_yes_button.click()
                time.sleep(2)
                element = driver.find_element(By.ID, value = "toast-container")
                message = element.text
                if "Successful" in message:
                    column_name = 'Transaction ID'
                    transaction_id = row[ column_name ]
                    data.at[ index, "Status" ] = "Mentioned transaction has been reversed"
                    print(transaction_id, "Reversed")
                    # If the "Details" button was found, it's a customer number
                    file_path = os.path.join(script_dir, "MRR.xlsx")
                    df = pd.read_excel(file_path)
                    row_number = index  # Assuming you want to start from the first row (index 0)
                    column_name = 'Status_Message'
                    df.at[ row_number, column_name ] = "Reversed"
                    df.to_excel(file_path, index = False, engine = 'openpyxl')
                    driver.refresh()
                else:
                    column_name = 'Transaction ID'
                    transaction_id = row[ column_name ]
                    data.at[ index, "Status" ] = "Not Reversed"
                    print(transaction_id, "Not Reversed")
                    driver.refresh()
            else:
                column_name = 'Transaction ID'
                transaction_id = row[column_name]
                data.at[ index, "Status" ] = "Mentioned transaction has been reversed already"
                print(transaction_id, "Mentioned transaction has been reversed already")
                driver.refresh()

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
