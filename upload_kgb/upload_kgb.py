# importing webdriver from selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import pandas as pd
import time
import os


def open_website():
    """Open EHRM website and return driver"""
    print("Opening EHRM website...")
    driver = webdriver.Chrome()

    # Fullscreen
    driver.maximize_window()

    url = "https://ehrm.pu.go.id/"
    driver.get(url)

    return driver


def click_enter_button(driver):
    """Find and click 'Masuk' button"""
    try:
        # Wait maximum 10 seconds for button to appear
        wait = WebDriverWait(driver, 10)

        # Find button with class btn btn-primary
        button = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button.btn.btn-primary"))
        )

        print("'Masuk' button found!")
        print(f"Text: {button.text}")

        # Click button
        button.click()
        print("Button clicked!")

        # Wait a few seconds after click
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error: {e}")
        print("Button not found or failed to click")
        input("Press Enter to continue after fixing the issue...")
        return False


def click_sso_login_button(driver):
    """Find and click 'Login SSO' button"""
    try:
        # Wait maximum 10 seconds for button to appear
        wait = WebDriverWait(driver, 10)

        # Find button with class btn btn-warning
        button = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button.btn.btn-warning"))
        )

        print("'Login SSO' button found!")
        print(f"Text: {button.text}")

        # Click button
        button.click()
        print("'Login SSO' button clicked!")

        # Wait a few seconds after click
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error: {e}")
        print("'Login SSO' button not found or failed to click")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_username(driver, username):
    """Fill username field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find username input field
        username_field = wait.until(
            EC.presence_of_element_located((By.ID, "UsernameId"))
        )

        print("Username field found!")
        username_field.clear()
        username_field.send_keys(username)
        print(f"Username '{username}' filled!")

        return True

    except Exception as e:
        print(f"Error filling username: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_password(driver, password):
    """Fill password field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find password input field
        password_field = wait.until(
            EC.presence_of_element_located((By.ID, "PasswordId"))
        )

        print("Password field found!")
        password_field.clear()
        password_field.send_keys(password)
        print("Password filled!")

        return True

    except Exception as e:
        print(f"Error filling password: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def click_login_button(driver):
    """Find and click Login button"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find submit button with text 'Login'
        login_button = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "button[type='submit'].btn.btn-primary"))
        )

        print("Login button found!")

        # Click button
        login_button.click()
        print("Login button clicked!")

        # Wait a few seconds after click
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error clicking login button: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def open_employee_service_page(driver):
    """Open Employee Service page (Layanan Kepegawaian Daftar Pegawai)"""
    try:
        url = "https://ehrm.pu.go.id/layanan-kepegawaian/daftar-pegawai"
        print(f"Opening URL: {url}")
        driver.get(url)

        # Wait for page to load
        time.sleep(3)

        print("Employee Service page has been opened!")
        return True

    except Exception as e:
        print(f"Error opening Employee Service page: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def read_all_employees_from_excel():
    """Read Excel file and get all employee data from PNS Satker sheet"""
    try:
        # Find Excel file in current directory
        excel_file = None
        for file in os.listdir('.'):
            if file.endswith('.xlsx') and 'kgb' in file.lower():
                excel_file = file
                break

        if not excel_file:
            print("Error: Excel file with 'kgb' in name not found!")
            return None

        print(f"Excel file found: {excel_file}")

        # Read "PNS Satker" sheet
        df = pd.read_excel(excel_file, sheet_name='PPPK Pusat')

        # Get all rows
        if len(df) == 0:
            print("Error: PNS Satker sheet is empty!")
            return None

        print(f"Total employees found: {len(df)}")
        return df

    except Exception as e:
        print(f"Error reading Excel: {e}")
        return None


def fill_nip(driver, nip):
    """Fill NIP field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find NIP input field using XPath
        nip_field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/div/div/div/div[1]/form/div[1]/div[1]/div/input'))
        )

        print("NIP field found!")
        nip_field.clear()
        nip_field.send_keys(nip)
        print(f"NIP '{nip}' filled!")

        return True

    except Exception as e:
        print(f"Error filling NIP: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def click_search_button(driver):
    """Find and click 'Search Data' button"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find button with text "Cari Data"
        button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/div/div/div/div[1]/form/div[3]/div[1]/button'))
        )

        print("'Search Data' button found!")

        # Click button
        button.click()
        print("'Search Data' button clicked!")

        # Wait a few seconds for search results
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error clicking search button: {e}")
        # Try alternative selector
        try:
            button = driver.find_element(
                By.XPATH, "//button[contains(., 'Cari Data')]")
            button.click()
            print("'Search Data' button clicked (via XPath)!")
            time.sleep(3)
            return True
        except:
            print(f"Failed to find search button")
            input("Press Enter to continue after fixing the issue...")
            return False


def click_view_employee_button(driver):
    """Find and click 'View Employee' button (Lihat Pegawai)"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find button using XPath
        button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="datatable-basic"]/tbody/tr/td[10]/div/div[1]/a/button'))
        )

        print("'View Employee' button found!")

        # Scroll to element before clicking
        driver.execute_script("arguments[0].scrollIntoView(true);", button)
        time.sleep(1)

        # Click button
        button.click()
        print("'View Employee' button clicked!")

        # Wait a few seconds for page to load
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error clicking view employee button: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def click_salary_button(driver):
    """Find and click 'Salary' button (Gaji)"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find button using XPath
        button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[3]/div/div/div/button[4]'))
        )

        print("'Salary' button found!")

        # Click button
        button.click()
        print("'Salary' button clicked!")

        # Wait a few seconds for page to load
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error clicking salary button: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def click_add_data_button(driver):
    """Find and click 'Add Data' button"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find button using XPath
        button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="gaji"]/div/div[1]/div/button'))
        )

        print("'Add Data' button found!")

        # Click button
        button.click()
        print("'Add Data' button clicked!")

        # Wait a few seconds for page to load
        time.sleep(3)

        return True

    except Exception as e:
        print(f"Error clicking add data button: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_salary_change_type(driver):
    """Fill 'Jenis Perubahan Gaji' dropdown with 'Kenaikan Gaji Berkala'"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and click the dropdown
        dropdown = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="vs8__combobox"]'))
        )

        print("Salary change type dropdown found!")
        dropdown.click()
        time.sleep(1)

        # Try to find option with scroll retry
        max_scroll_attempts = 5
        scroll_attempt = 0
        option = None

        while scroll_attempt < max_scroll_attempts:
            try:
                option = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='vs8__option-1']"))
                )
                print(f"Option found!")
                break
            except:
                scroll_attempt += 1
                if scroll_attempt < max_scroll_attempts:
                    print(
                        f"Option not found, scrolling dropdown container... (attempt {scroll_attempt}/{max_scroll_attempts})")
                    # Scroll the dropdown list itself, not the window
                    try:
                        option_list = driver.find_element(
                            By.XPATH, "//*[@role='listbox']")
                        driver.execute_script(
                            "arguments[0].scrollBy(0, 150);", option_list)
                        time.sleep(0.5)
                    except:
                        print("Could not find listbox, trying alternative scroll...")
                        driver.execute_script("window.scrollBy(0, 100);")
                        time.sleep(0.5)
                else:
                    print(f"Max scroll attempts reached for salary type")

        if option:
            option.click()
            print("'Kenaikan Gaji Berkala' selected!")
        else:
            print("Error: Could not find salary type option")
            return False

        return True

    except Exception as e:
        print(f"Error filling salary change type: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_sk_number(driver, sk_number):
    """Fill 'No. SK' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div/div/input'))
        )
        print("SK number field found!")
        field.clear()
        field.send_keys(str(sk_number))
        print(f"SK number '{sk_number}' filled!")

        return True

    except Exception as e:
        print(f"Error filling SK number: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_sk_date(driver, sk_date_xpath):
    """Fill 'Tanggal SK' field using datepicker calendar"""
    try:
        from datetime import datetime
        import os
        wait = WebDriverWait(driver, 10)

        # Find the datepicker input field
        date_field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[3]/div[2]/div/div/div/input'))
        )

        print("SK date field found!")

        # Click to open datepicker (ensure widget initialized)
        date_field.click()
        time.sleep(1)

    # Find the date button in the datepicker
        date_button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, sk_date_xpath))
        )

        print("SK date button found!")

        # Click to open datepicker (ensure widget initialized)
        date_button.click()
        time.sleep(1)

        return True

    except Exception as e:
        print(f"Error filling SK date: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_golongan(driver, golongan):
    """Fill 'Golongan' field"""
    try:
        wait = WebDriverWait(driver, 10)
        dropdown = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="vs9__combobox"]'))
        )

        print("Golongan dropdown found!")
        dropdown.click()
        time.sleep(1)
        print(golongan)

        # Try to find option with scroll retry
        max_scroll_attempts = 5
        scroll_attempt = 0
        option = None

        while scroll_attempt < max_scroll_attempts:
            try:
                option = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//li[@role='option' and text()='{golongan}']"))
                )
                print(f"Golongan option found!")
                break
            except:
                scroll_attempt += 1
                if scroll_attempt < max_scroll_attempts:
                    print(
                        f"Option not found, scrolling dropdown container... (attempt {scroll_attempt}/{max_scroll_attempts})")
                    # Scroll the dropdown list itself, not the window

                    try:
                        option_list = driver.find_element(
                            By.XPATH, '//*[@id="vs9__listbox"]')
                        driver.execute_script(
                            "arguments[0].scrollBy(0, 200);", option_list)
                        time.sleep(0.5)
                    except:
                        print("Could not find listbox, trying alternative scroll...")
                        # Fallback: try scrolling any visible dropdown container
                        driver.execute_script("window.scrollBy(0, 100);")
                        time.sleep(0.5)
                else:
                    print(
                        f"Max scroll attempts reached for golongan '{golongan}'")

        if option:
            option.click()
            print(f"Golongan '{golongan}' selected!")
        else:
            print(f"Error: Could not find golongan '{golongan}' in dropdown")
            return False

        return True

    except Exception as e:
        print(f"Error filling golongan: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_work_year_old(driver, tahun_masa_kerja_lama):
    """Fill 'Tahun Masa Kerja Lama' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[5]/div[2]/div/div/input'))
        )

        print("Old work year field found!")
        field.clear()
        field.send_keys(str(tahun_masa_kerja_lama))
        print(f"Old work year '{tahun_masa_kerja_lama}' filled!")

        return True

    except Exception as e:
        print(f"Error filling old work year: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_work_month_old(driver, bulan_masa_kerja):
    """Fill 'Bulan Masa Kerja Lama' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[6]/div[2]/div/div/input'))
        )

        print("Old work month field found!")
        field.clear()
        field.send_keys(str(bulan_masa_kerja))
        print(f"Old work month '{bulan_masa_kerja}' filled!")

        return True

    except Exception as e:
        print(f"Error filling old work month: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_work_year_new(driver, tahun_masa_kerja_baru):
    """Fill 'Tahun Masa Kerja Baru' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[7]/div[2]/div/div/input'))
        )

        print("New work year field found!")
        field.clear()
        field.send_keys(str(tahun_masa_kerja_baru))
        print(f"New work year '{tahun_masa_kerja_baru}' filled!")

        return True

    except Exception as e:
        print(f"Error filling new work year: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_work_month_new(driver, bulan_masa_kerja_baru):
    """Fill 'Bulan Masa Kerja Baru' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[8]/div[2]/div/div/input'))
        )

        print("New work month field found!")
        field.clear()
        field.send_keys(str(bulan_masa_kerja_baru))
        print(f"New work month '{bulan_masa_kerja_baru}' filled!")

        return True

    except Exception as e:
        print(f"Error filling new work month: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def upload_kgb_file(driver, employee_name):
    """Upload KGB file based on employee name"""
    try:
        import os
        import glob
        wait = WebDriverWait(driver, 10)

        # Construct file search pattern
        kgb_folder = "kgb"
        search_pattern = os.path.join(
            kgb_folder, "**", f"*{employee_name}*.pdf")

        # Search for the file
        matching_files = glob.glob(search_pattern, recursive=True)

        if not matching_files:
            print(f"No KGB file found for employee: {employee_name}")
            input("Press Enter to continue after placing the correct file...")
            return False

        # Use the first matching file
        file_path = matching_files[0]
        print(f"Found KGB file: {file_path}")

        # Find the file upload element
        file_input = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//input[@type="file"]'))
        )

        print("File upload input found!")

        # Upload the file
        file_input.send_keys(os.path.abspath(file_path))
        print(f"KGB file uploaded: {file_path}")

        return True

    except Exception as e:
        print(f"Error uploading KGB file: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_pejabat_penandatangan(driver, pejabat_penandatangan):
    """Fill 'Pejabat Penandatangan' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field
        field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[11]/div[2]/div/div/input'))
        )

        print("Pejabat penandatangan field found!")
        field.clear()
        field.send_keys(str(pejabat_penandatangan))
        print(f"Pejabat penandatangan '{pejabat_penandatangan}' filled!")

        return True

    except Exception as e:
        print(f"Error filling pejabat penandatangan: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_effective_date(driver, effective_date_xpath, next_month_clicks=1):
    """Fill 'Tanggal mulai' field using datepicker calendar"""
    try:
        from datetime import datetime
        import os
        wait = WebDriverWait(driver, 10)

        # Find the datepicker input field
        date_field = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[10]/div[2]/div/div[1]/div/input'))
        )

        print("Effective date field found!")

        # Click to open datepicker (ensure widget initialized)
        date_field.click()
        time.sleep(1)

        # Find next month button and click specified number of times
        next_month_button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[2]/div/div[1]/div[10]/div[2]/div/div[2]/div/div[2]/div/div[1]/div/div[1]/button[2]'))
        )

        print("Next month button found!")

        # Click next month button the specified number of times
        for i in range(next_month_clicks):
            print(
                f"Clicking next month button (click {i+1}/{next_month_clicks})")
            next_month_button.click()
            time.sleep(1)

    # Find the date button in the datepicker
        date_button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, effective_date_xpath))
        )

        print("Effective date button found!")

        # Click to open datepicker (ensure widget initialized)
        date_button.click()
        time.sleep(1)

        return True

    except Exception as e:
        print(f"Error filling effective date: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def fill_salary_new(driver, gaji_pokok_baru):
    """Fill 'Gaji Pokok Baru' field"""
    try:
        wait = WebDriverWait(driver, 10)

        # Find and fill the field - use reliable selectors only
        field = None
        selectors_to_try = [
            # Exact placeholder match
            '//input[@placeholder="Gaji Pokok Baru"]',
            '//input[contains(@class, "v-money3")]',  # Unique class identifier
            '//input[@id="219"]',  # Original ID as fallback
        ]

        for selector in selectors_to_try:
            try:
                field = wait.until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                print(f"New salary field found using selector: {selector}")
                break
            except:
                continue

        if not field:
            raise Exception(
                "Could not find new salary field with any selector")

        print("New salary field found!")
        field.clear()
        field.send_keys(str(gaji_pokok_baru))
        print(f"New salary '{gaji_pokok_baru}' filled!")

        return True

    except Exception as e:
        print(f"Error filling new salary: {e}")
        input("Press Enter to continue after fixing the issue...")
        return False


def verify_submission_success(driver):
    """Verify that the submission was successful by checking for the expected element with username"""
    try:
        # Get expected username from environment
        expected_username = os.getenv('EXPECTED_USERNAME')
        if not expected_username:
            print("Warning: EXPECTED_USERNAME not found in .env file")
            input("Press Enter to continue without username verification...")
            return True
        
        wait = WebDriverWait(driver, 10)
        
        # Check for the specific element that indicates successful submission
        element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="datatable-basic"]/tbody/tr[1]/td[8]'))
        )
        
        element_text = element.text
        print(f"Element content found: {element_text}")
        print(f"Checking for username: {expected_username}")
        
        # Check if the element contains the expected username
        if expected_username in element_text:
            print(f"✓ Submission verification successful - {expected_username} found in element!")
            return True
        else:
            print(f"✗ {expected_username} not found in element content")
            print(f"Expected username '{expected_username}' but found different content")
            input("Press Enter to continue after checking the submission...")
            return False
        
    except Exception as e:
        print(f"Error verifying submission: {e}")
        print("Expected element not found - submission may have failed")
        input("Press Enter to continue after checking the submission status...")
        return False


def main():
    """Main function"""
    # Load environment variables from .env file
    load_dotenv()

    # Get username and password from .env
    username = os.getenv('USERNAME_EHRM')
    password = os.getenv('PASSWORD_EHRM')

    # Optional SK date override and source xpath from .env
    sk_date_xpath = os.getenv('SK_DATE_XPATH')
    print("PASTIKAN SK_DATE DI .env SUDAH BENAR!")
    print(f"SKDATE_SOURCE_XPATH (env): {sk_date_xpath}")

    # Optional Effective date override and source xpath from .env
    effective_date_xpath = os.getenv('EFFECTIVE_DATE_XPATH')
    next_month_clicks = os.getenv('EFFECTIVE_DATE_NEXT_MONTH_CLICKS', '1')
    print("PASTIKAN EFFECTIVE_DATE DI .env SUDAH BENAR!")
    print(f"EFFECTIVE_DATE_XPATH (env): {effective_date_xpath}")
    print(f"EFFECTIVE_DATE_NEXT_MONTH_CLICKS (env): {next_month_clicks}")

    if not username or not password:
        print("Error: USERNAME_EHRM or PASSWORD_EHRM not found in .env file")
        print("Make sure .env file exists with format:")
        print("USERNAME_EHRM=your_username")
        print("PASSWORD_EHRM=your_password")
        return

    driver = None

    try:
        employees_df = read_all_employees_from_excel()

        driver = open_website()
        click_enter_button(driver)
        click_sso_login_button(driver)

        # Fill username and password from .env
        fill_username(driver, username)
        fill_password(driver, password)

        # Click login button
        click_login_button(driver)

        # Open Employee Service page
        open_employee_service_page(driver)

        # Read all employees from Excel at the beginning

        if employees_df is not None:
            # Loop through each employee
            for index, row in employees_df.iterrows():
                try:
                    print(f"\n{'='*60}")
                    print(
                        f"Processing employee {index + 1} of {len(employees_df)}")
                    print(f"{'='*60}")

                    # Convert row to dictionary
                    employee_data = row.to_dict()

                    # Get NIP from employee data
                    nip = employee_data.get('nip')

                    if nip:
                        print(f"NIP: {nip}")

                        # Fill NIP
                        fill_nip(driver, nip)

                        # Click search button
                        click_search_button(driver)

                        # Click view employee button
                        click_view_employee_button(driver)

                        # Click salary button
                        click_salary_button(driver)

                        # Click add data button
                        click_add_data_button(driver)

                        # Scroll down to see form
                        driver.execute_script("window.scrollBy(0, 500);")
                        time.sleep(1)

                        # Fill salary change type
                        fill_salary_change_type(driver)

                        # Get data from employee_data to fill form
                        sk_number = employee_data.get('nomor_surat')
                        fill_sk_number(driver, sk_number)

                        fill_sk_date(driver, sk_date_xpath)

                        # Scroll down to see form
                        driver.execute_script("window.scrollBy(0, 300);")
                        time.sleep(1)
                        golongan = employee_data.get('golongan')
                        fill_golongan(driver, golongan)

                        tahun_masa_kerja_lama = employee_data.get(
                            'acuan_terakhir_masa_kerja_tahun')
                        fill_work_year_old(driver, tahun_masa_kerja_lama)

                        bulan_masa_kerja_lama = employee_data.get(
                            'acuan_terakhir_masa_kerja_bulan')
                        fill_work_month_old(driver, bulan_masa_kerja_lama)

                        tahun_masa_kerja_baru = employee_data.get(
                            'baru_masa_kerja_tahun')
                        fill_work_year_new(driver, tahun_masa_kerja_baru)

                        bulan_masa_kerja_baru = employee_data.get(
                            'baru_masa_kerja_bulan')
                        fill_work_month_new(driver, bulan_masa_kerja_baru)

                        gaji_pokok_baru = employee_data.get('baru_gaji')
                        fill_salary_new(driver, gaji_pokok_baru)

                        # Scroll down to see form
                        driver.execute_script("window.scrollBy(0, 500);")
                        time.sleep(1)

                        next_month_clicks = int(next_month_clicks)
                        fill_effective_date(
                            driver, effective_date_xpath, next_month_clicks)

                        pejabat_penandatangan = employee_data.get(
                            'baru_jabatan_pejabat_penandatangan')
                        fill_pejabat_penandatangan(
                            driver, pejabat_penandatangan)

                        # Upload KGB file
                        nama = employee_data.get('nama')
                        upload_kgb_file(driver, nama)
                        
                        time.sleep(5)

                        # input("Press Enter to continue to next employee...")

                        confirm_button = driver.find_element(
                            By.XPATH, '//*[@id="app"]/div/main/div/section/section/div[2]/div/div[2]/div[3]/div/button[2]')
                        confirm_button.click()

                        time.sleep(5)
                        
                        # Verify submission was successful before continuing
                        if verify_submission_success(driver):
                            print("✓ Employee submission verified successfully")
                        else:
                            print("✗ Employee submission verification failed")
                            input("Press Enter to continue to next employee anyway...")
                        
                        print(
                            f"Successfully processed employee with NIP: {nip}")
                        # Open Employee Service page
                        open_employee_service_page(driver)
                    else:
                        print("Warning: NIP column not found in employee data")

                except Exception as e:
                    print(f"Error processing employee {index + 1}: {e}")
                    continue

        # Keep the browser open
        input("Press Enter to close browser...")

    except Exception as e:
        print(f"Error in main: {e}")

    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()
