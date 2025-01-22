# Retry mechanism for operations that may fail
import time
from venv import logger


def retry(func, max_retries=3, wait_time=2 * 60):
    attempt = 1
    while attempt <= max_retries:
        try:
            return func()  # Call the function and return its result
        except Exception as e:
            logger.error(f"Error during attempt {attempt}: {str(e)}")
            if attempt < max_retries:
                logger.info(f"Retrying... (Attempt {attempt + 1} of {max_retries})")
            else:
                logger.warning("Maximum Retry attempts reached. Waiting for 2 minutes before retrying.")
                time.sleep(wait_time)
                attempt += 1

                # After max retries, retry indefinitely
            logger.info("Retrying indefinitely after max retries...")
            while True:
                try:
                    return func()  # Try again after the wait time and return its result
                except Exception as e:
                    logger.error(f"Error during retry: {str(e)}")
                    logger.info("Retrying after 2 minutes...")
                    time.sleep(wait_time)