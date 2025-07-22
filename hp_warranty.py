import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, NoSuchFrameException
from webdriver_manager.chrome import ChromeDriverManager
import os

# --- Script Configuration ---
serial_number_file = "C:/Users/uhfnb/OneDrive - Barnes Group Inc/Dokumente/Code/Python/serials.xlsx"
output_excel_file = "warranty_results.xlsx"
website_url = "https://support.hp.com/us-en/check-warranty"

serial_number_column_name = "Serial Number"
product_number_column_name = "Model Number"
input_device_name_column_name = "Device Name"

output_new_device_name_header = "Device Name"
output_new_status_header = "Status"
output_new_end_date_header = "End Date"
output_new_start_date_header = "Start Date"

output_columns_order = [
    output_new_device_name_header,
    "Serial Number",
    output_new_status_header,
    output_new_start_date_header,
    output_new_end_date_header
]


# --- Function to Setup Chrome WebDriver ---
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--log-level=3")
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    print("Setting up Chrome WebDriver...")
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        # Using implicit wait for general element presence,
        # but explicit waits are used for specific interactions.
        driver.implicitly_wait(15)
        return driver
    except ValueError as e:
        print("\n---")
        print("ERROR: Could not start Chrome. Please ensure Google Chrome is installed on your system.")
        print(f"DETAILS: {e}")
        print("---\n")
        return None

# --- Function to Handle Cookie Consent ---
def handle_cookie_consent(driver, wait_time=10):
    local_wait = WebDriverWait(driver, wait_time)

    print("    - Attempting to handle cookie banner...")
    try:
        driver.switch_to.default_content()
    except WebDriverException as e:
        print(f"    - Warning: Could not switch to default content initially: {e}")
        pass

    try:
        # Try to accept the cookie button directly (common case)
        cookie_button = local_wait.until(EC.element_to_be_clickable((By.XPATH,
            "//button[contains(.,'Accept') or contains(.,'Alle Cookies akzeptieren') or contains(.,'Accept All') or contains(.,'Ich stimme zu')]"
        )))
        cookie_button.click()
        print("    - Accepted cookie policy directly.")
        # No time.sleep here, next wait will handle the page state
        return True
    except (TimeoutException, NoSuchElementException):
        pass # Button not found directly, proceed to check for iframe
    except Exception as e:
        print(f"    - Unexpected error trying to click cookie button directly: {type(e).__name__}: {e}")

    try:
        # Check for cookie iframe
        iframe_element = local_wait.until(EC.presence_of_element_located((By.ID, "onetrust-pc-sdk")))
        try:
            driver.switch_to.frame(iframe_element)
            print("    - Switched to cookie consent iframe.")
            cookie_button_in_iframe = local_wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
            cookie_button_in_iframe.click()
            print("    - Accepted cookie policy within iframe.")
            return True
        except (TimeoutException, NoSuchElementException) as nested_e:
            print(f"    - Timeout/No such element *inside* cookie iframe: {type(nested_e).__name__}. Cookie banner not handled.")
            return False
        except NoSuchFrameException as nested_e:
            print(f"    - Warning: Cookie iframe disappeared or became invalid before interaction: {nested_e.msg}. Cookie banner not handled.")
            return False
        except WebDriverException as nested_e:
            print(f"    - WebDriver Error *inside* cookie iframe handling block: {type(nested_e).__name__}: {nested_e}. Cookie banner not handled.")
            return False
        except Exception as nested_e:
            print(f"    - Unforeseen error *inside* cookie iframe handling block: {type(nested_e).__name__}: {nested_e}. Cookie banner not handled.")
            return False
        finally:
            try:
                driver.switch_to.default_content()
                print("    - Switched back to main content from iframe block.")
            except WebDriverException as e:
                print(f"    - Warning: Could not switch to default content in iframe finally block: {e}")
                pass
    except (TimeoutException, NoSuchElementException) as e:
        print(f"    - Cookie iframe element not found (Timeout/No Such Element): {type(e).__name__}. Cookie banner not handled.")
    except Exception as e:
        print(f"    - Unforeseen error trying to locate cookie iframe element: {type(e).__name__}: {e}. Cookie banner not handled.")

    print("    - Cookie banner not handled.")
    return False


# --- Main Script Execution ---
if __name__ == "__main__":
    driver = None
    df_output = pd.DataFrame(columns=output_columns_order)

    # Resume Logic: Load existing results
    processed_serials = set()
    if os.path.exists(output_excel_file):
        print(f"'{output_excel_file}' found. Checking for previously processed serial numbers...")
        try:
            df_existing_results = pd.read_excel(output_excel_file)
            if "Serial Number" in df_existing_results.columns:
                processed_serials = set(df_existing_results["Serial Number"].astype(str).tolist())
                df_output = df_existing_results.copy()
                print(f"Found {len(processed_serials)} serial numbers already processed.")
            else:
                print(f"Warning: 'Serial Number' column not found in '{output_excel_file}'. Starting fresh.")
        except Exception as e:
            print(f"Error reading existing results file '{output_excel_file}': {e}. Starting fresh.")
    else:
        print(f"'{output_excel_file}' not found. Starting a new results file.")


    # Read Input Serial Numbers
    if not os.path.exists(serial_number_file):
        print(f"\nERROR: The file '{serial_number_file}' was not found.")
        print("Please ensure the Excel file exists at the specified path.")
        exit()

    try:
        df_input_serials = pd.read_excel(serial_number_file)
        required_columns = [serial_number_column_name, product_number_column_name, input_device_name_column_name]
        for col in required_columns:
            if col not in df_input_serials.columns:
                print(f"\nERROR: The Excel input file must contain a column named '{col}'.")
                print("Please check the column headers in your Excel file or update the COLUMN_NAME variables in the script.")
                exit()

        all_serials_data = df_input_serials[[serial_number_column_name, product_number_column_name, input_device_name_column_name]].astype(str).values.tolist()

    except Exception as e:
        print(f"\nERROR: Could not read the Excel file '{serial_number_file}'.")
        print(f"DETAILS: {e}")
        print("Please ensure the file is a valid Excel file and not open in another program.")
        exit()

    if not all_serials_data:
        print(f"The input columns in '{serial_number_file}' are empty. No serial numbers to check.")
        exit()

    # Filter out already processed serial numbers for resuming
    serials_to_process = []
    for sn, pn, device_name in all_serials_data:
        if sn not in processed_serials:
            serials_to_process.append((sn, pn, device_name))

    if not serials_to_process:
        print("All serial numbers already processed or no new serial numbers to check. Exiting.")
        if driver:
            driver.quit()
        exit()

    print(f"Found {len(all_serials_data)} total serial number(s) in input file.")
    print(f"Will process {len(serials_to_process)} new/remaining serial number(s).")


    driver = setup_driver()
    if driver is None:
        exit()

    # Define a default wait for general page interactions
    wait = WebDriverWait(driver, 60)

    total_serials_overall = len(all_serials_data)
    start_index_overall = total_serials_overall - len(serials_to_process)

    # Define the width of the progress bar
    overall_bar_length = 50
    scraping_bar_length = 20 # New constant for the inner bar

    for idx_in_subset, (sn, pn_from_excel, device_name_from_input) in enumerate(serials_to_process):
        current_overall_index = start_index_overall + idx_in_subset + 1
        overall_progress_percentage = (current_overall_index / total_serials_overall) * 100

        # Calculate the number of filled characters for the overall bar
        filled_length_overall = int(overall_bar_length * overall_progress_percentage // 100)
        overall_bar = '█' * filled_length_overall + '-' * (overall_bar_length - filled_length_overall)

        # Print the overall progress bar line
        print(f"\n--- [{overall_bar}] {overall_progress_percentage:.1f}% ({current_overall_index}/{total_serials_overall}) - Serial Number: {sn} ---")

        current_serial_result_dict = {
            output_new_device_name_header: device_name_from_input,
            "Serial Number": sn,
            output_new_status_header: "Processing Error",
            output_new_start_date_header: "N/A",
            output_new_end_date_header: "N/A",
            "Product Number Used": "N/A"
        }

        try:
            driver.get(website_url)
            # Removed time.sleep(3) here. WebDriverWait below will wait for the input box.

            cookie_handled = handle_cookie_consent(driver)
            if not cookie_handled:
                print("    - Warning: Cookie banner could not be handled. This might affect subsequent steps.")
            # Removed time.sleep(2) here. The next wait will handle the page state after cookie interaction.

            input_box = wait.until(EC.element_to_be_clickable((By.ID, "inputtextpfinder")))
            input_box.clear()
            input_box.send_keys(sn)
            print(f"    - Entered serial number: {sn}")
            # Removed time.sleep(1) here. The click action below will trigger a page load/update.

            submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "FindMyProduct")))
            driver.execute_script("arguments[0].click();", submit_btn)
            print("    - Submitted serial number. Checking for product number prompt or direct results...")

            try:
                # Use a specific wait for either the product number input or the info section
                element_found = WebDriverWait(driver, 30).until(
                    EC.any_of(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[formcontrolname='productNumber']")),
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "div.info-section"))
                    ),
                    message="Timed out waiting for either product number prompt or info section after SN submission."
                )

                if element_found and element_found.get_attribute('formcontrolname') == "productNumber":
                    print("    - 'Product number' input field detected. Prompt is present.")
                    pn_input_box = element_found

                    if not pn_from_excel or pn_from_excel.lower() == 'nan':
                        print(f"    - Warning: Product/Model number for '{sn}' is empty/NaN in Excel. Cannot fill prompt.")
                    else:
                        pn_input_box.clear()
                        pn_input_box.send_keys(pn_from_excel)
                        print(f"    - Entered product number: {pn_from_excel}")
                        current_serial_result_dict["Product Number Used"] = pn_from_excel

                        submit_pn_btn = wait.until(EC.element_to_be_clickable((By.ID, "FindMyProductNumber")))
                        driver.execute_script("arguments[0].click();", submit_pn_btn)
                        print("    - Re-submitted with product number.")

                        # Wait for the info section after product number submission
                        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.info-section")))
                        print("    - Final info section found after product number submission.")

                elif element_found and "info-section" in element_found.get_attribute('class'):
                    print("    - 'Product number' prompt not detected. Direct results page loaded.")

            except TimeoutException as e:
                print(f"    - Timeout: Neither Product Number prompt nor main info section appeared after serial submission: {e.msg}")
                current_serial_result_dict[output_new_status_header] = "Navigation Timeout"
                current_serial_result_dict[output_new_end_date_header] = "Error"
                current_serial_result_dict[output_new_start_date_header] = "Error"
            except Exception as e:
                print(f"    - ERROR: An unexpected error occurred during dynamic element detection: {type(e).__name__}: {e}")
                current_serial_result_dict[output_new_status_header] = "Dynamic Detection Error"
                current_serial_result_dict[output_new_end_date_header] = "Error"
                current_serial_result_dict[output_new_start_date_header] = "Error"

            if current_serial_result_dict[output_new_status_header] == "Processing Error":
                try:
                    info_section_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.info-section")))
                    print(f"    - Main info section confirmed for scraping.")

                    warranty_status_scraped = "Not Found"
                    warranty_end_date_scraped = "Not Found"
                    warranty_start_date_scraped = "Not Found"

                    try:
                        info_items = info_section_element.find_elements(By.CSS_SELECTOR, "div.info-item")
                        total_info_items = len(info_items)

                        for item_idx, item in enumerate(info_items):
                            scraping_progress_percentage = ((item_idx + 1) / total_info_items) * 100
                            
                            # Calculate the number of filled characters for the scraping bar
                            filled_length_scraping = int(scraping_bar_length * scraping_progress_percentage // 100)
                            scraping_bar = '█' * filled_length_scraping + '-' * (scraping_bar_length - filled_length_scraping)

                            # Print the scraping progress bar line
                            print(f"        - Scraping data: [{scraping_bar}] {scraping_progress_percentage:.1f}% ({item_idx+1}/{total_info_items})", end='\r')
                            
                            try:
                                label_elem = item.find_element(By.CSS_SELECTOR, "div.label")
                                text_elem = item.find_element(By.CSS_SELECTOR, "div.text")
                                label = label_elem.text.strip()

                                p_tags = text_elem.find_elements(By.TAG_NAME, "p")
                                if p_tags:
                                    value = "\n".join([p.text.strip() for p in p_tags if p.text.strip()])
                                else:
                                    value = text_elem.text.strip()

                                if label == "End date":
                                    warranty_end_date_scraped = value
                                elif label == "Status":
                                    warranty_status_scraped = value
                                elif label == "Start date":
                                    warranty_start_date_scraped = value
                            except NoSuchElementException:
                                continue
                            except Exception as item_e:
                                print(f"\n    - Error parsing info-item: {item_e}")
                        print("\n") # Newline after the scraping progress line to clear the '\r' effect

                        print(f"    - Success: Scraped details for '{current_serial_result_dict[output_new_device_name_header]}'.")
                        print(f"        - Warranty Status: {warranty_status_scraped}")
                        print(f"        - Warranty Start Date: {warranty_start_date_scraped}")
                        print(f"        - Warranty End Date: {warranty_end_date_scraped}")
                        
                        current_serial_result_dict[output_new_status_header] = warranty_status_scraped
                        current_serial_result_dict[output_new_end_date_header] = warranty_end_date_scraped
                        current_serial_result_dict[output_new_start_date_header] = warranty_start_date_scraped

                    except NoSuchElementException as e:
                        current_serial_result_dict[output_new_status_header] = "Scraping Error"
                        current_serial_result_dict[output_new_end_date_header] = "Error"
                        current_serial_result_dict[output_new_start_date_header] = "Error"
                        print(f"    - Error: Could not parse page for warranty details (info-section content missing): {e}")
                    except TimeoutException as e:
                        current_serial_result_dict[output_new_status_header] = "Scraping Timeout"
                        current_serial_result_dict[output_new_end_date_header] = "Error"
                        current_serial_result_dict[output_new_start_date_header] = "Error"
                        print(f"    - ERROR: Timeout during scraping process. Elements not found after result section visibility confirmed: {e}")
                    except Exception as e:
                        current_serial_result_dict[output_new_status_header] = "Unhandled Scraping Error"
                        current_serial_result_dict[output_new_end_date_header] = "Error"
                        current_serial_result_dict[output_new_start_date_header] = "Error"
                        print(f"    - UNEXPECTED ERROR during scraping: {e}")
                except TimeoutException:
                    print(f"    - ERROR: A timeout occurred waiting for info section for SN {sn}. Page content not ready.")
                    current_serial_result_dict[output_new_status_header] = "Info Section Timeout"
                    current_serial_result_dict[output_new_end_date_header] = "Error"
                    current_serial_result_dict[output_new_start_date_header] = "Error"
                except Exception as e:
                    print(f"    - UNEXPECTED ERROR getting info section for SN {sn}: {e}")
                    current_serial_result_dict[output_new_status_header] = "Info Section Error"
                    current_serial_result_dict[output_new_end_date_header] = "Error"
                    current_serial_result_dict[output_new_start_date_header] = "Error"

        except TimeoutException:
            print(f"    - ERROR: A timeout occurred for SN {sn}. The page may have failed to load or a specific element was not found in time.")
            current_serial_result_dict[output_new_status_header] = "Global Timeout"
            current_serial_result_dict[output_new_end_date_header] = "Error"
            current_serial_result_dict[output_new_start_date_header] = "Error"
        except WebDriverException as e:
            print(f"    - CRITICAL ERROR: WebDriver error encountered for SN {sn}: {e}")
            print("    - The WebDriver session may have been lost. Attempting to re-establish the driver.")
            current_serial_result_dict[output_new_status_header] = "WebDriver Session Lost"
            current_serial_result_dict[output_new_end_date_header] = "Error"
            current_serial_result_dict[output_new_start_date_header] = "Error"
            if driver:
                driver.quit()
            driver = setup_driver()
            if driver is None:
                print("    - Failed to re-establish WebDriver. Exiting script.")
                break
            else:
                wait = WebDriverWait(driver, 60)
                # Removed time.sleep(5) here. The next iteration will try to navigate.
        except Exception as e:
            print(f"    - UNEXPECTED GLOBAL ERROR for SN {sn}: An unhandled error occurred: {e}")
            current_serial_result_dict[output_new_status_header] = "Unhandled Script Error"
            current_serial_result_dict[output_new_end_date_header] = "Error"
            current_serial_result_dict[output_new_start_date_header] = "Error"

        new_row_df = pd.DataFrame([current_serial_result_dict])
        if "Product Number Used" not in output_columns_order and "Product Number Used" in new_row_df.columns:
            new_row_df = new_row_df.drop(columns=["Product Number Used"])
        new_row_df = new_row_df[output_columns_order]

        df_output = pd.concat([df_output, new_row_df], ignore_index=True)

        try:
            df_output.to_excel(output_excel_file, index=False)
            print(f"    - Results saved incrementally to '{output_excel_file}'")
        except Exception as e:
            print(f"    - WARNING: Could not save incremental results to '{output_excel_file}'. Is the file open? Details: {e}")

        # Removed time.sleep(2) here. The next loop iteration will implicitly wait for the next page load.

    print("\n--------------------------------------------------")
    print("All serial numbers have been processed (or script terminated due to critical error).")

    if not df_output.empty:
        try:
            df_output.to_excel(output_excel_file, index=False)
            print(f"Final results saved successfully to: '{output_excel_file}'")
        except Exception as e:
            print(f"ERROR: Could not perform final save to '{output_excel_file}'. Please ensure the file is not open: {e}")
    else:
        print("No results were generated for saving.")

    if driver:
        driver.quit()
    print("WebDriver closed. Script finished.")
    print("--------------------------------------------------")
