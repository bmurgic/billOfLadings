import pyautogui
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

# Open a new instance of Safari with the webdriver
browser = webdriver.Safari()

# Create ActionsChain object
a = ActionChains(browser)

# Notes to add in details of BOL
IN_OFF = "PPW in office"


def process_bols(filename, note=IN_OFF):

    # Prepare the browser and workflow
    prepare_bol_workflow()

    # Open the file given
    bols_f = open(filename, 'r')
    bols = [line.strip() for line in bols_f]
    bols_f.close()

    # Perform BOLs workflow
    updated_successfully, errored_out = perform_bol_workflow(bols)

    print("Complete!")
    # Add all successfully updated BOLs to the completed file
    if len(updated_successfully) > 0:
        print("Writing BOLs to 'bols_completed.txt'...", end='')
        while True:
            try:
                complete_f = open("bols_completed.txt", "a")
                for bol in updated_successfully:
                    complete_f.write(bol + "\n")
                complete_f.close()
                break
            except Exception as e:

                # Alert the user that it failed to save the completed BOLs
                print("Failed to open bols_completed.txt. Exception raised: " + str(e))

                # Get user input if should retry saving
                try_again = pyautogui.confirm(text="Failed to open bols_completed.txt. Exception raised: " + str(e),
                                              title="Bill of Ladings Workflow",
                                              buttons=['OK', 'Cancel'])

                if try_again in 'Cancel':
                    break

        print("Success!")

    # Add all successfully updated BOLs to the completed file
    if len(errored_out) > 0:
        print("Writing BOLs to 'bols_error.txt'...", end='')
        while True:
            try:
                errored_f = open("bols_error.txt", "a")
                for bol in errored_out:
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

        print("Success!")
        print("There were " + str(len(updated_successfully)) + " updated successfully!", end='')
        print("There were " + str(len(errored_out)) + " with errors!")
        input("Press enter to close: ")

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
                password_elem.send_keys(str(password))  # Enter the password
                password_elem.submit()  # Log in

                try:
                    # Once logged in, wait for the home page to load
                    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@href='/Home/Search']")))

                    break
                except Exception as e:
                    password = password = input("Incorrect password. Try again: ")

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

        try:
            WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@id='tbSearchText']")))
        except Exception as e:
            print("Search input not found! Exception raised: ", str(e))
            browser.close()
            exit()

        try:  # Mark checkbox to search for exact BOL number
            check_box_input = browser.find_element(By.XPATH, "//div[@id='cbIsEqual']/input")
            browser.execute_script("$(arguments[0]).click();", check_box_input)
            print("Success!")
        except (TimeoutException, NoSuchElementException):
            print("Checkbox not found!")

        try:
            WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.ID, "_uiq_ft")))

            close_button_modal = browser.find_element(By.XPATH, "//div[@id='_uiq_ft']//button[@class='uiq_close']")
            close_button_modal.click()
        except (TimeoutException, NoSuchElementException):
            pass
        except Exception as e:
            print("Could not close dialog! Exception raised: " + str(e))
            input("Press enter to close: ")
            browser.close()
            exit()


def search_bol(bol):

    # Start search
    try:
        # Find search box
        search_box_input = browser.find_element(By.XPATH, "//div[@id='tbSearchText']//input")  # Get search box input
        browser.execute_script("$(arguments[0]).click();", search_box_input)  # Click search box
        search_box_input.clear()  # Clear any text inside input
        search_box_input.send_keys(bol)  # Enter in the BOL number
    except Exception as e:
        return "Could not enter " + bol + " into search box. Exception raised: " + str(e)

    try:  # Click the search button
        browser.find_element(By.ID, "bSearch").click()
        return bol + " searched successfully! "
    except Exception as e:
        return "Could not search " + bol + ". Exception raised: " + str(e) + " "


def update_bol_notes(bol, note=IN_OFF):

    try:  # Open BOLs details and enter it is in office
        WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, "//td[text()='" + str(bol) + "']")))
    except (TimeoutException, NoSuchElementException):
        return "BOL could not be found or does not exist!"
    except Exception as e:
        return "BOL could not be found! Error raised: " + str(e)

    try:
        # Perform the addition of notes to the BOLs details
        browser.find_element(By.CLASS_NAME, "invoice-detail-popup").click()  # Open the BOLs details popup
    except Exception as e:
        return "Could not open BOL details. Exception raised: " + str(e)

    try:  # Wait for the popup to finish opening
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='noteTextBoxInPopup']//textarea")))
        notes = browser.find_element(By.XPATH, "//div[@id='noteTextBoxInPopup']//textarea")  # Find the element for entering notes
        notes.send_keys(note)  # Enter that the PPW is in the office

        add_button = browser.find_element(By.ID, "btnAddInvoiceNote")  # Find the element to add details
        add_button.click()  # Click add button
    except Exception as e:
        return "Could not update details for BOL. Exception raised: " + str(e)

    try:  # Close the details popup


        # Close the popup
        browser.find_element(By.XPATH, "//button[text()='Close']").click()

        return bol + " updated!"
    except Exception as e:
        return "Could not close BOL details. Exception raised: " + str(e)


def perform_bol_workflow(bols):

    updated_successfully = []
    errored_out = []

    while len(bols):

        # Checks if the search page is open, otherwise open it
        go_to_search_page()

        # Once bol has been ran over, then remove it from the list
        bol = bols.pop(0)

        # Search for the BOL
        searched = search_bol(bol)

        # If it was able to successfully search for the BOL, then update its details
        if 'not' not in searched:
            updated = update_bol_notes(bol)

            # If it was able to successfully update the BOLs details, then add it to the list of completed BOLs
            if 'not' not in updated:
                updated_successfully.append(bol)
            else:
                errored_out.append(searched + updated)  # If it failed to update the BOL, add it to the list of errors
        else:
            errored_out.append(searched)  # If it failed to search the BOL, then add it to the list of errors

    return updated_successfully, errored_out


process_bols("bols.txt")