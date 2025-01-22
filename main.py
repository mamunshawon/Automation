from venv import logger
import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


def main():
    # script_dir = os.path.dirname(os.path.abspath(__file__))
    # input_file_path = os.path.join(script_dir, "DKYC.xlsx")
    # base_output_dir = os.path.join(script_dir, "DKYC_Output")
    # os.makedirs(base_output_dir, exist_ok=True)
    # Load data
    data = load_excel_data(input_file_path)
    logger.info(f"Total rows to process: {len(data)}")
    # Initialize WebDriver
    driver = webdriver.Chrome()
    driver.maximize_window()
    # Load existing results if available
    # output_file_path = os.path.join(base_output_dir, "DKYC_Results.xlsx")
    # if os.path.exists(output_file_path):
    #    existing_results = pd.read_excel(output_file_path, dtype={'NID': str})
    # else:
    #    existing_results = pd.DataFrame()  # Initialize as empty DataFrame
    results = []
    try:
        # EC portal login
        try:
            ec_portal_login(driver, "dkyc.automation2", "N@gad1@34")
            time.sleep(5)
            # Nagad portal login
            nagad_portal_login(driver, "sysops.automation@gmail.com", "Nagad@202404")
            logger.info("Both logins successful, proceeding with the next steps.")
            time.sleep(15)
        except Exception as e:
            logger.error(f"Login failed: {str(e)}")
            return False  # Indicating failure
        for index, row in data.iterrows():
            try:
                nid = str(row['NID']).strip()
                full_dkyc = str(row['DKYC']).strip()
                dkyc_number = full_dkyc.replace("D-", "")  # Numeric-only DKYC (e.g., 1138262452)
                status = str(row['Status']).strip()
                # Update status to "In-Progress"
                # data.at[index, 'Status'] = 'In-Progress'
                # data.to_excel(input_file_path, index=False)  # Save progress
                if status == "Ready to Process":
                    # Process DKYC row EC Results
                    ec_data = process_dkyc_row(driver, row, base_output_path)
                    print(ec_data)
                    if ec_data == "Invalid NID or DOB":
                        input_status = "Invalid NID or DOB"

                    elif ec_data != "Invalid NID or DOB":
                        ec_data["Full DKYC"] = full_dkyc
                        nid_folder = ec_data.get("Folder Path", os.path.join(base_output_path, nid))
                        process_nagad_portal(driver, ec_data["Full DKYC"], base_output_path,
                                             nid)  # Process Nagad portal
                        source_file = f"./DKYC_Output/{nid}/{dkyc_number}_{nid}_NGD.jpg"
                        target_file = f"./DKYC_Output/{nid}/{nid}_EC.jpg"

                        face_matches = compare_faces(source_file, target_file)

                        # face_matches_data = face_matches["Face Match"]
                        # Define the dictionary
                        # face_matches_data = {
                        #     "Similarity": '',
                        #     "Face match": ''
                        # }
                        # Extract the 'Face match' value
                        face_matches_data = face_matches.get("Face match", None)
                        print(face_matches_data)
                        # Create a WebDriverWait instance
                        nagad_portal(driver, ec_data["Full DKYC"], base_output_path, nid)  # Process Nagad portal
                        wait = WebDriverWait(driver, 10)  # Define wait here
                        # Ensure that 'wait' is defined as a WebDriverWait instance
                        status_column_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                           "2]/div/div/div/app-kyc-list/section/div/div/div["
                                                           "2]/div/div/app-common-table-advanced/table/tbody/tr/td["
                                                           "7]/span/span")
                        status_column = wait.until(EC.visibility_of_element_located(status_column_locator))

                        # Check if the status is "CS ACCEPTED" or "CS REJECTED"
                        status_value = status_column.text.strip()
                        print(status_value)
                        # Prepend "Already" if the status is "CS ACCEPTED" or "CS REJECTED"
                        if status_value in ["CS ACCEPTED", "CS REJECTED"]:
                            status_value = f"Already {status_value}"
                            print(status_value)
                            updated_status = {"KYC_Status": status_value}
                            result = {**ec_data, **face_matches, **updated_status}
                            results.append(result)
                            results_df = pd.DataFrame(results)
                            results_df.to_excel(output_file_path, index=False)
                            logger.info(f"Results saved to {output_file_path}")

                        elif status_value == "CS RECEIVED":
                            process_face_match_action(driver, full_dkyc, face_matches_data)
                            nagad_portal(driver, ec_data["Full DKYC"], base_output_path, nid)  # Process Nagad portal
                            wait = WebDriverWait(driver, 10)  # Define wait here
                            # Ensure that 'wait' is defined as a WebDriverWait instance
                            status_column_locator = (By.XPATH, "/html/body/app-root/app-full-layout/div/div["
                                                               "2]/div/div/div/app-kyc-list/section/div/div/div["
                                                               "2]/div/div/app-common-table-advanced/table/tbody/tr/td["
                                                               "7]/span/span")
                            status_column = wait.until(EC.visibility_of_element_located(status_column_locator))
                            status_value = status_column.text.strip()
                            updated_status = {"KYC_Status": status_value}
                            result = {**ec_data, **face_matches, **updated_status}
                            results.append(result)
                            results_df = pd.DataFrame(results)
                            results_df.to_excel(output_file_path, index=False)
                            logger.info(f"Results saved to {output_file_path}")

                        input_status = "Done"

                    try:
                        # Update status to "Done" after successful processing
                        data.at[index, 'Status'] = input_status
                        data.to_excel(input_file_path, index=False)

                    except Exception:
                        data.at[index, 'Status'] = 'Could not Process'
                        data.to_excel(input_file_path, index=False)

                delete_all_in_directory(base_output_path)
            # else:

            except Exception as e:
                # Capture the error and update status with the reason where it failed
                error_message = f"Error at step: {str(e)}"
                pattern = r"Message: (.+)"
                match = re.search(pattern, error_message)
                if match:
                    parsed_message = match.group(1)
                    logger.error(parsed_message)
                else:
                    logger.error("error_message")

                data.at[index, 'Status'] = f"Failed: {error_message}"  # Write the error message to the Status column
                data.to_excel(input_file_path, index=False)  # Save progress

    finally:
        # Quit the driver at the end of all operations
        driver.quit()


if __name__ == "__main__":
    main()
