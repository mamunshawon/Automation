import time

from pandas import read_excel
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

# Open the Excel file and read the data into a pandas dataframe
data = read_excel("Book.xlsx")

username = "uatdemo18@gmail.com"
password = "N@gad1234"

driver: WebDriver = webdriver.Chrome(executable_path='./driver/chromedriver.exe')
driver.get('https://systest.mynagad.com:20020/ui/system/#/home')
driver.maximize_window()
time.sleep(3)

driver.find_element(By.ID, "username").send_keys(username)
driver.find_element(By.ID, "password").send_keys(password)

try:
    driver.find_element(By.ID, "login_button").click()
    # Find the validation message element
    validation_messages = driver.find_elements_by_id("validation-message")

    # # Check if any validation messages were found
    # if len(validation_messages) > 0:
    #     # Print the text of the first validation message
    #     print(validation_messages[0].text)
    # else:
    #     print("No validation message found on the page")
except NoSuchElementException:
    print("Either the button or the validation message was not found on the page")

for index, row in data.iterrows():
    driver.get('https://sys.mynagad.com:20020/ui/system/#/bill-pay-management/biller-service-detail/174368')
    time.sleep(2)

    try:
        driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                               "2]/div/div/div/app-biller-service-detail/section/div"
                                               "/div[1]/div/div[2]/div/div/form/div/button[3]").click()
    except NoSuchElementException:
        print("Element not found on the page")

    time.sleep(2)
    Service_Name = driver.find_element(by=By.XPATH, value="//*[@id='serviceName']")
    Service_Name_Ba = driver.find_element(By.ID, "serviceNameBn")
    Service_Number = driver.find_element(By.ID, "serviceNumber")
    Service_Name.send_keys(row["First Name"])
    time.sleep(2)
    Service_Name_Ba.send_keys(row["Last Name"])
    time.sleep(2)
    Service_Number.send_keys(row["Service Number"])
    time.sleep(2)
    try:
        driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                               "2]/div/div/div/app-biller-service/section/div["
                                               "2]/div/div/div/div/form/div/div["
                                               "6]/div/div/div/span/i").click()
    except NoSuchElementException:
        print("Element not found on page")

    try:
        Enter_merchant_number = driver.find_element(by=By.XPATH, value="/html/body/ngb-modal-window/div/div/app-common"
                                                                       "-search-merchant/div/div[2]/div["
                                                                       "1]/div/div/form/div[1]/div/div[2]/div/div/input")
        Enter_merchant_number.send_keys('0')
        Enter_merchant_number.send_keys(row["Merchant Number"])
    except NoSuchElementException:
        print("Element Not Found")
        time.sleep(2)

    click_search_button = driver.find_element(by=By.XPATH, value="/html/body/ngb-modal-window/div/div/app-common"
                                                                 "-search-merchant/div/div[2]/div[1]/div/div/form/div["
                                                                 "2]/div/button")
    click_search_button.click()
    time.sleep(2)

    click_bullet_button = driver.find_element(by=By.XPATH, value="/html/body/ngb-modal-window/div/div/app-common"
                                                                 "-search-merchant/div/div[2]/div["
                                                                 "2]/div/table/tbody/tr/td[4]/input")
    click_bullet_button.click()

    click_select_button = driver.find_element(by=By.XPATH, value="/html/body/ngb-modal-window/div/div/app-common"
                                                                 "-search-merchant/div/div[3]/span[2]")
    click_select_button.click()
    time.sleep(2)

    click_edit_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                               "2]/div/div/div/app-biller-service/section/div["
                                                               "2]/div/div/div/div/form/div/div["
                                                               "27]/div/table/tbody/tr[4]/td[5]/button")
    click_edit_button.click()
    time.sleep(2)

    # ScrollPage #
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    # ScrollPage #
    driver.execute_script("window.scroll(0, 0);")
    time.sleep(2)

    driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                           "2]/div/div/div/app-biller-service/section/div["
                                           "2]/div/div/div/div/form/div/div[26]/div[2]/div/div/div/div[1]/div["
                                           "2]/div/div/div/input").clear()

    driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                           "2]/div/div/div/app-biller-service/section/div["
                                           "2]/div/div/div/div/form/div/div[26]/div[2]/div/div/div/div[1]/div["
                                           "2]/div/div/div/input").send_keys('0')

    driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                           "2]/div/div/div/app-biller-service/section/div["
                                           "2]/div/div/div/div/form/div/div[26]/div[2]/div/div/div/div[1]/div["
                                           "2]/div/div/div/input").send_keys(row["Merchant Number"])
    time.sleep(2)

    click_update_button = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-full-layout/div/div["
                                                                 "2]/div/div/div/app-biller-service/section/div["
                                                                 "2]/div/div/div/div/form/div/div[26]/div["
                                                                 "8]/div/button")
    click_update_button.click()
    time.sleep(2)

    try:
        click_register_button = driver.find_element(by=By.XPATH, value="//*[@id='horizontal-form-layouts']/div["
                                                                       "2]/div/div/div/div/form/div/div[31]/div/button")
        click_register_button.click()
        # validation_messages = driver.find_elements_by_id("validation-message")
        # # Check if any validation messages were found
        # if len(validation_messages) > 0:
        #     # Print the text of the first validation message
        #     print(validation_messages[0].text)
        # else:
        #     print("No validation message found on the page")
    except NoSuchElementException:
        print("Element Not Found")
        time.sleep(2)

    # get and save output data
    # output = driver.find_element_by_id("output_field").Result
    # data.at[i, "output"] = output

# write data to excel file
# data.to_excel('Book.xls', index=False)
# Close the browser
driver.quit()

# driver.find_element(By.ID, "mobileNumber1").send_keys(mobileNumber1)
# ime.sleep(1)
