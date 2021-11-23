import pyautogui
import gspread
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# Open a new instance of Safari with the webdriver
# browser = webdriver.Safari()
browser = webdriver.Chrome()

# Set size of browser
browser.set_window_position(0, 0)
browser.set_window_size(1600, 900)

# Create ActionsChain object
a = ActionChains(browser)

# Create client to access Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = gspread.authorize(credentials)
worksheet = gc.open("BOLs History").sheet1
todays_date = datetime.today().strftime('%m/%d/%Y')

# Notes to add in details of BOL
IN_OFF = "PPW in office"


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)


def write_to_history_spreadsheet(driver, bol, success):
    # First find the next open spot
    next_row = next_available_row(worksheet)

    # Then write to history
    worksheet.update_acell("A{}".format(next_row), driver)
    worksheet.update_acell("B{}".format(next_row), bol)
    worksheet.update_acell("C{}".format(next_row), todays_date)
    worksheet.update_acell("D{}".format(next_row), success)

    print(" Added to history spreadsheet!")


def process_bols(filename):

    # Prepare the browser and workflow
    prepare_bol_workflow()

    # Open the filename given
    bols_f = open(filename, 'r')
    bols = [line.strip() for line in bols_f]
    bols_f.close()

    # Perform BOLs workflow
    perform_bol_workflow(bols)

    print("Complete!")

    # Alert the user that the workflow is complete
    browser.close()
    exit()


def prepare_bol_workflow():

    print("Preparing workflow...", end='')
    go_to_otr_website()
    login()
    go_to_search_page()
    print("Workflow ready!\nProcessing BOLs")


def go_to_otr_website():

    # Check if search page is open
    current_url = browser.current_url

    if 'Login' in current_url:
        print("Login page open!")
    elif 'Dashboard' in current_url:
        print("Dashboard open!")
    elif 'Search' in current_url:
        print("Search page already open!")
    else:
        print("Opening OTR's website...", end='')

        # Open OTR's website
        browser.get("https://otrcapitalportal.com/")


def login():

    current_url = browser.current_url

    if 'Auth' in current_url:
        try:  # Attempt to log in, otherwise already logged in

            print("Logging in...", end='')

            password = input("Enter password: ")

            if password == "None":
                browser.close()
                exit()

            # Waits until it finds the username field
            WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.NAME, 'UserName')))

            # If the username field is found then log in
            email_elem = browser.find_element(By.NAME, "UserName")  # Get username field
            email_elem.send_keys("brandon@traveloko.com")  # Enter in email
            password_elem = browser.find_element(By.NAME, "Password")  # Get password field

            while True:
                password_elem.clear()
                password_elem.send_keys(str(password))  # Enter the password
                password_elem.submit()  # Log in

                try:
                    # Once logged in, wait for the home page to load
                    WebDriverWait(browser, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//a[@href='/Home/Search']")))

                    break
                except Exception as e:
                    password = input("Incorrect password. Try again: ")

                    if password == "None":
                        browser.close()
                        exit()

            print("Successfully logged in!")

        except Exception as e:
            print("Unable to log in! Exception raised: ", str(e))
            browser.close()
            exit()


def go_to_search_page():

    # Check if search page is open
    current_url = browser.current_url

    if 'Search' not in current_url:
        print("Opening search page...", end='')

        # Go to the search page
        browser.get("https://otrcapitalportal.com/Home/Search")
        tried = 0

        while True:
            try:
                WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@id='tbSearchText']")))
                break
            except Exception as e:

                if tried > 5:
                    print("Too many tries...restart BOL Workflow!")
                    browser.close()
                    exit()

                print("Search input not found! Retrying...")
                go_to_search_page()
                tried += 1

        try:  # Mark checkbox to search for exact BOL number
            check_box_input = browser.find_element(By.XPATH, "//div[@id='cbIsEqual']/input")
            browser.execute_script("$(arguments[0]).click();", check_box_input)
            print("Success!")
        except (TimeoutException, NoSuchElementException):
            print("Checkbox not found!")

        try:
            WebDriverWait(browser, 1).until(EC.presence_of_element_located((By.ID, "_uiq_ft")))

            close_button_modal = browser.find_element(By.XPATH, "//div[@id='_uiq_ft']//button[@class='uiq_close']")
            close_button_modal.click()
        except (TimeoutException, NoSuchElementException):
            pass
        except Exception as e:
            print("Could not close dialog! Exception raised: " + str(e))
            input("Press enter to close: ")
            browser.close()
            exit()

        a.click(browser.find_element(By.ID, "ddlClient")).perform()
        for _ in range(3):
            a.send_keys(Keys.TAB).perform()
            sleep(0.3)
        a.send_keys(Keys.ENTER).perform()


def search_bol(bol):

    # Start search
    # Find search box
    search_input = browser.find_element(By.XPATH, "//div[@id='tbSearchText']//input")
    # a.click(browser.find_element(By.XPATH, "//div[@id='tbSearchText']")).perform()
    search_input.clear()
    search_input.send_keys(bol)
    print("Searching for " + bol + "...", end='')

    try:  # Click the search button
        search_button = browser.find_element(By.ID, "bSearch")
        search_button.click()
        # browser.execute_script("$(arguments[0]).click();", search_button)  # Click search box

        WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='dgSearchBoard']//td[text()='" + bol + "']")))

        print("found...", end='')
        return bol + " searched successfully! "
    except Exception as e:
        print("couldn't find BOL!", end='')
        return "Could not search " + bol + ". Exception raised: " + str(e) + " "


def update_bol_notes(bol, note=IN_OFF):

    try:  # Open BOLs details
        WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "invoice-detail-popup")))
        popup_link = browser.find_element(By.CLASS_NAME, "popup-link")
        # a.click(popup_link).perform()
        browser.execute_script("arguments[0].click();", popup_link)

        # Wait for the details popup finish loading
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='FilesGrid']//span[@class='show-file']")))
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='noteTextBoxInPopup']//textarea")))
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='invoiceInfoGrid']/div/div/div/table/tbody/tr/td")))
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@id='notesGrid']/div/div/div/div/div/div/table/tbody/tr/td")))
        notes_input = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='noteTextBoxInPopup']//textarea")))

        print("opening details...", end='')
    except Exception as e:
        print("couldn't open details!", end='')
        return "Could not open BOL details. Exception raised: " + str(e)

    try:  # Add notes to BOL details

        # Write in office
        notes_input.send_keys(note)
        print("wrote in office...", end='')

        # Add the note
        add_button = browser.find_element(By.ID, "btnAddInvoiceNote")  # Find the element to add details
        add_button.click()

        # Check if it was added before closing
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='notesGrid']//td[contains(text(), 'PPW in office')]")))
        print("successfully updated notes...", end='')

    except Exception as e:
        print("failed to update notes!", end='')
        return "Could not update details for BOL. Exception raised: " + str(e)

    try:  # Close the details popup
        close_button = browser.find_element(By.XPATH, "//button[text()='Close']")
        browser.execute_script("arguments[0].click();", close_button)
        sleep(1.5)
        return bol + " updated!"
    except Exception as e:
        print("couldn't close notes!", end='')
        return "Could not close BOL details. Exception raised: " + str(e)


def write_to_file(filename, bol):

    while True:
        try:
            errored_f = open(filename, "a")
            errored_f.write(bol + "\n")
            errored_f.close()
            break
        except Exception as e:

            # Alert the user that it failed to save the errored BOLs
            print("Failed to open bols_error.txt. Exception raised: " + str(e))

            # Get user input if should retry saving
            try_again = pyautogui.confirm(text="Failed to open bols_error.txt. Exception raised: " + str(e),
                                          title="Bill of Ladings Workflow",
                                          buttons=['OK', 'Cancel'])

            if try_again in 'Cancel':
                break

    print(" Added to '" + filename + "!")


def perform_bol_workflow(bols):
    driver = 0
    bol = 0

    while len(bols):

        # Checks if the search page is open, otherwise open it
        go_to_search_page()

        # Once bol has been ran over, then remove it from the list
        line = bols.pop(0)
        if len(line) == 3:
            driver = line
            bol = bols.pop(0)
        elif len(line) >= 5:
            bol = line

        try:
            # Search for the BOL
            searched = search_bol(bol)

            # If it was able to successfully search for the BOL, then update its details
            if 'not' not in searched:
                updated = update_bol_notes(bol)

                # If it was able to successfully update the BOLs details, then add it to the file of completed BOLs
                if 'not' not in updated:
                    # write_to_file("bols_completed.txt", bol)
                    write_to_history_spreadsheet(driver, bol, "YES")
                else:
                    # write_to_file("bols_error.txt", bol)
                    write_to_history_spreadsheet(driver, bol, "NO")
            else:
                # write_to_file("bols_error.txt", bol)
                write_to_history_spreadsheet(driver, bol, "NO")

        except Exception as e:
            print("Was not able to complete entire BOL Workflow. Exception was raised: " + str(e))
            browser.close()
            exit()


process_bols("bols.txt")
