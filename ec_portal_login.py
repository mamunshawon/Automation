import time

import logger
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


# EC portal login
def ec_portal_login(driver, username, password, max_retries=3, wait_time=2 * 60):
    login_url = 'https://prportal.nidw.gov.bd/partner-portal/login'
    for attempt in range(1, max_retries + 1):
        try:
            driver.get(login_url)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'login-username')))
            driver.find_element(By.ID, 'login-username').send_keys(username)
            driver.find_element(By.ID, 'login-password').send_keys(password)
            driver.find_element(By.ID, "login-button").click()
            logger.info("EC portal login successful.")
            return  # Exit the function if login is successful
        except WebDriverException as e:
            # Log a concise error message without the full stack trace
            if 'net::ERR_CONNECTION_TIMED_OUT' in str(e):
                logger.error("EC Portal Unreachable. Connection timed out.")
            else:
                logger.error("Error during EC Portal login.")

            if attempt < max_retries:
                logger.info(f"Retrying EC portal login... (Attempt {attempt + 1} of {max_retries})")
            else:
                logger.warning("Maximum retry attempts reached for EC portal. Waiting for 2 minutes before retrying.")
                time.sleep(wait_time)
                # Retry after waiting for 2 minutes
                return ec_portal_login(driver, username, password, max_retries, wait_time)

    # If all retries and wait fail, raise an exception
    raise Exception("Failed to login EC Portal after multiple attempts or server unreachable.")
