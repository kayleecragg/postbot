import os
import time
import zipfile
import subprocess
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, ElementClickInterceptedException
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import dearpygui.dearpygui as dpg


def run_game(competition, timer, folder, spreadsheet_id, range_name):
    SPREADSHEET_ID = spreadsheet_id
    RANGE_NAME = range_name

    tennis = competition
    new_folder_name = 'Post Images'
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    download_default_directory = os.path.join(downloads_folder, new_folder_name)
    os.makedirs(download_default_directory, exist_ok=True)

    # Google Sheets API setup
    creds = Credentials.from_service_account_file(f'{folder}/tilebot1-a31d2f8e1c87.json')
    service = build('sheets', 'v4', credentials=creds)

    #logos image input
    logos_base_path = f'{folder}/Flags'

    # Set up the Chrome options
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_default_directory,
    "download.prompt_for_download": False,  # To disable the download prompt
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
    })

    # Create a browser instance
    browser = webdriver.Chrome(options=chrome_options)

    # Function to clear input field
    def clear_input_field(element):
        # Clear using Selenium's clear method
        element.clear()
        # If the element is not cleared, use JavaScript to clear it
        browser.execute_script("arguments[0].value = '';", element)
        # Check and retry if the input is not cleared
        for i in range(3):
            if element.get_attribute('value'):
                element.clear()
                browser.execute_script("arguments[0].value = '';", element)
                time.sleep(1)
            else:
                break

    # Call the Sheets API to read data
    sheet_service = service.spreadsheets()
    result = sheet_service.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    # Process and iterate over Google Sheets data
    if not values:
        print('No data found.')
    else:
        for row in values[1:]:
            # Ensure row has enough elements
            row += [None] * (14 - len(row))

            # Unpack the row into variables
            input1_value, input2_value, round_value, datetime_value, gender_value, sd_value, court_value, rank1_value, rank2_value, event_data, countries1_value, countries2_value, firstname1_value, firstname2_value = row

            if input1_value is None:  # Check for end of data
                break

            browser.get('https://thelivecms.prod.streamco.cloud/tile-creator/')
            overlay_script = """
            var overlay = document.createElement('div');
            overlay.id = 'selenium-overlay';
            overlay.style.position = 'fixed';
            overlay.style.top = '0';
            overlay.style.left = '0';
            overlay.style.width = '100%';
            overlay.style.height = '100%';
            overlay.style.backgroundColor = 'rgba(0,0,0,0.5)';
            overlay.style.zIndex = '10000';
            overlay.style.pointerEvents = 'none'; // Allows Selenium interactions
            document.body.appendChild(overlay);
            """
            browser.execute_script(overlay_script)
            # live & upcoming

            live = WebDriverWait(browser, 2, 0.2).until(
                EC.presence_of_element_located(
                    (By.XPATH,  "//*[contains(text(), 'Tennis Tiles')]")
                )
            )
            live.click();

            # home & away

            home = WebDriverWait(browser, 2, 0.2).until(
                EC.presence_of_element_located(
                    (By.XPATH,  "//*[contains(text(), 'Tennis Replay')]")
                )
            )
            home.click()



            # select sport + competition

            select1 = WebDriverWait(browser, 2, 0.2).until(
                EC.presence_of_element_located(
                    (By.ID, "select-sport-type-id")
                )
            )
            select1 = Select(select1)
            select1.select_by_visible_text('Tennis')


            select2 = WebDriverWait(browser, 2, 0.2).until(
                EC.presence_of_element_located(
                    (By.ID, "select-competition-id")
                )
            )
            select2 = Select(select2)
            if tennis == 'Australian Open':
                if sd_value == 'Singles':
                    select2.select_by_visible_text('TENNIS: Australian Open Singles')
                else:
                    select2.select_by_visible_text('TENNIS: Australian Open Extra Comps')
            if tennis == 'Roland Garros':
                if sd_value == 'Singles':
                    select2.select_by_visible_text('TENNIS: Roland Garros Singles')
                else:
                    select2.select_by_visible_text('TENNIS: Roland Garros Doubles')
            if tennis == 'Wimbledon':
                if sd_value == 'Singles':
                    select2.select_by_visible_text('TENNIS: Wimbledon Singles')
                else:
                    select2.select_by_visible_text('TENNIS: Wimbledon Doubles')
            if tennis == 'US Open':
                if sd_value =='Singles':
                    select2.select_by_visible_text('TENNIS: US Open Singles')
                else:
                    select2.select_by_visible_text('TENNIS: US Open Doubles')
                
            
            if sd_value == 'Doubles':
                fourplayers_checkbox = WebDriverWait(browser, 2, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,  "//label[text()='4 Players']/preceding-sibling::input")
                    )
                )
                fourplayers_checkbox.click()
                    

            if '/' in input1_value:
                # Split the string into two parts at the slash
                name1_value, name2_value = input1_value.split('/')

                # name 1

                name1 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='home1']")
                    )
                )
                clear_input_field(name1)

                name1 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='home1']")
                    )
                )
                name1.send_keys(name1_value)

                # name 2
                name2 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='home2']")
                    )
                )
                clear_input_field(name2)

                name2 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='home2']")
                    )
                )
                name2.send_keys(name2_value)

                # Now name1 and name2 hold the separate names
                # You can use name1 and name2 as needed in your code
            else:
                time.sleep(1)
                input1 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.ID, "downshift-0-input")
                    )
                )
                clear_input_field(input1)
                input1 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.ID, "downshift-0-input")
                    )
                )

                input1.send_keys(input1_value)

                # team 2

                input2 = browser.find_element(By.ID,"downshift-1-input")
                clear_input_field(input2)
                input2.send_keys(input2_value)

        
            if '/' in input2_value:
                # Split the string into two parts at the slash
                name3_value, name4_value = input2_value.split('/')

                name3 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='away1']")
                    )
                )
                clear_input_field(name3)

                name3 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='away1']")
                    )
                )
                name3.send_keys(name3_value)

                name4 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='away2']")
                    )
                )
                clear_input_field(name4)

                name4 = WebDriverWait(browser, 10, 0.2).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[value='away2']")
                    )
                )
                name4.send_keys(name4_value)

            # rank 1

            if sd_value == 'Singles':
                if rank1_value:  # This will be False if rank1_value is blank or None
                    rank1 = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #1 Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank1.send_keys(rank1_value)

            # rank 2
            if sd_value == 'Singles':
                if rank2_value:  # This will be False if rank1_value is blank or None
                    rank2 = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #2 Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank2.send_keys(rank2_value)

            # Doubles Rank
                    
            # Rank 1A + 1B
            if sd_value == 'Doubles':
                if rank1_value:  # This will be False if rank1_value is blank or None
                    rank1A = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #1A Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank1A.send_keys(rank1_value)

                    rank1B = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #1B Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank1B.send_keys(rank1_value)

            # Rank 2A + 2B
            if sd_value == 'Doubles':
                if rank2_value:  # This will be False if rank1_value is blank or None
                    rank2A = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #2A Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank2A.send_keys(rank2_value)

                    rank2B = WebDriverWait(browser, 10, 0.2).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, "//label[contains(text(), 'Opponent #2B Rank')]/following-sibling::input[@type='text']")
                        )
                    )
                    rank2B.send_keys(rank2_value)

            if '/' in countries1_value:
                # Split the string into two parts at the slash
                country1_value, country2_value = countries1_value.split('/')

                # opponent 1A flag
                file_input_opponent1A = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #1A Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent1A = os.path.join(logos_base_path, f"{country1_value}.png")
                file_input_opponent1A.send_keys(file_path_opponent1A)

                # opponent 1B flag
                file_input_opponent1B = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #1B Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent1B = os.path.join(logos_base_path, f"{country2_value}.png")
                file_input_opponent1B.send_keys(file_path_opponent1B)


            else:
                # opponent 1 flag
                file_input_opponent1 = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #1 Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent1 = os.path.join(logos_base_path, f"{countries1_value}.png")
                file_input_opponent1.send_keys(file_path_opponent1)

                # opponent 2 flag
                file_input_opponent2 = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #2 Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent2 = os.path.join(logos_base_path, f"{countries2_value}.png")
                file_input_opponent2.send_keys(file_path_opponent2)

            if '/' in countries2_value:
                # Split the string into two parts at the slash
                country3_value, country4_value = countries2_value.split('/')

                # opponent 2A flag
                file_input_opponent2A = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #2A Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent2A = os.path.join(logos_base_path, f"{country3_value}.png")
                file_input_opponent2A.send_keys(file_path_opponent2A)

                # opponent 2B flag
                file_input_opponent2B = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//label[contains(text(), 'Opponent #2B Flag')]/following-sibling::input[@type='file']")
                    )
                )
                file_path_opponent2B = os.path.join(logos_base_path, f"{country4_value}.png")
                file_input_opponent2B.send_keys(file_path_opponent2B)


            # delete event long

            event_long = WebDriverWait(browser, 10, 0.2).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[value='Event Long']")
                )
            )
            time.sleep(1)

            clear_input_field(event_long)

            time.sleep(1)

            event_long.send_keys(f"{gender_value} {sd_value}")
            
            event_short = WebDriverWait(browser, 10, 0.2).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[value='Event Short']")
                )
            )
            clear_input_field(event_short)

            if "Junior" in gender_value or "Wheelchair" in gender_value:
                event_short.send_keys(f"{gender_value}")
            else:
                event_short.send_keys(f"{gender_value} {sd_value}")


            # round 

            round_long = WebDriverWait(browser, 10, 0.2).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[value='Round Long']")
                )
            )
            clear_input_field(round_long)


            round_long.send_keys(round_value)
            
            round_short = WebDriverWait(browser, 10, 0.2).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[value='Round Short']")
                )
            )
            clear_input_field(round_short)

            if "Junior" in gender_value or "Wheelchair" in gender_value:
                round_short.send_keys(f"{sd_value} {round_value}")
            else:
                round_short.send_keys(round_value)
            
            time.sleep(2)

            remove_overlay_script = """
            var overlay = document.getElementById('selenium-overlay');
            if (overlay) {
                overlay.parentNode.removeChild(overlay);
            }
            """
            browser.execute_script(remove_overlay_script)
            
            time.sleep(timer)

            # downloading 

            try:
                # Wait for the button to be clickable by both class and text, with a timeout of 10 seconds
                save_as_zip_button = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-warning') and contains(text(), 'Save as .zip')]"))
                )

                time.sleep(2)

                
                # Instead of clicking the button with Selenium's click function, we use JavaScript to click
                browser.execute_script("arguments[0].click();", save_as_zip_button)
                print("Button clicked successfully via JavaScript.")
            except TimeoutException:
                print("Button was not found within the given time.")
            except NoSuchElementException:
                print("Button could not be found.")
            except Exception as e:
                print(f"An error occurred: {e}")

            time.sleep(5)

            download_directory = download_default_directory
            zip_files = sorted([f for f in os.listdir(download_directory) if f.endswith('.zip')], key=lambda f: os.path.getctime(os.path.join(download_directory, f)), reverse=True)

            if sd_value == 'Singles':
                if zip_files:
                    zip_file_name = zip_files[0]
                    zip_file_path = os.path.join(download_directory, zip_file_name)

                    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                        zip_ref.extractall(download_directory)
                        extracted_files = zip_ref.namelist()

                    if extracted_files:
                        new_singles_folder_path = os.path.join(download_directory, f"{input1_value} v {input2_value} {sd_value}")
                        if not os.path.exists(new_singles_folder_path):
                            os.mkdir(new_singles_folder_path)

                        print(f"Moving files to {new_singles_folder_path}")

                        for file in extracted_files:
                            original_file_path = os.path.join(download_directory, file)
                            new_file_path = os.path.join(new_singles_folder_path, file)
                            print(f"Moving {original_file_path} to {new_file_path}")
                            shutil.move(original_file_path, new_file_path)

                        subprocess.run(["open", "-R", new_singles_folder_path])
                        subprocess.run(["osascript", "-e", 'tell application "Finder" to close window 1'])


                    else:
                        print("No files found in the zip.")
                else:
                    print("No zip files found in the download directory.")

                os.remove(zip_file_path)
            
            else:
                name1_value, name2_value = input1_value.split('/')
                name3_value, name4_value = input2_value.split('/')
                if zip_files:
                    zip_file_name = zip_files[0]
                    zip_file_path = os.path.join(download_directory, zip_file_name)

                    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                        zip_ref.extractall(download_directory)
                        extracted_files = zip_ref.namelist()

                    if extracted_files:
                        new_doubles_folder_path = os.path.join(download_directory, f"{name1_value} v {name3_value} {sd_value}")
                        if not os.path.exists(new_doubles_folder_path):
                            os.mkdir(new_doubles_folder_path)

                        print(f"Moving files to {new_doubles_folder_path}")

                        for file in extracted_files:
                            original_file_path = os.path.join(download_directory, file)
                            new_file_path = os.path.join(new_doubles_folder_path, file)
                            print(f"Moving {original_file_path} to {new_file_path}")
                            shutil.move(original_file_path, new_file_path)

                        subprocess.run(["open", "-R", new_doubles_folder_path])
                        subprocess.run(["osascript", "-e", 'tell application "Finder" to close window 1'])

                    else:
                        print("No files found in the zip.")
                else:
                    print("No zip files found in the download directory.")

                os.remove(zip_file_path)

    browser.quit()
