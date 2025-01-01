import time
import pandas as pd
import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import winsound
from oauth2client.service_account import ServiceAccountCredentials

# Google Sheets setup
def load_indebted_shops_from_sheet(sheet_url):
    # Define the scope
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # Add credentials.json file path (download this from Google Cloud Console after enabling Sheets API)
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

    # Authorize the client
    client = gspread.authorize(creds)

    # Open the Google Sheet
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)  # Assuming data is in the first sheet

    # Extract data from the 'Business Name' column (Column A)
    shop_names = worksheet.col_values(1)  # Column A is indexed as 1

    # Convert all shop names to lowercase and strip spaces
    shop_names = [str(name).strip().lower() for name in shop_names if name]
    return shop_names

# Main function to monitor the list input
def monitor_shop_input():
    # Load the indebted shops from Google Sheet
    SHEET_URL = "https://docs.google.com/spreadsheets/d/13EOKkBNrfFFOQBoJF8UxVMCszQVpyntcBL0zNqBWPPc/edit?gid=1956153930#gid=1956153930"
    indebted_shops = load_indebted_shops_from_sheet(SHEET_URL)

    # Connect to an existing Chrome session using the remote debugging port
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"  # Connect to the debugging port
    # Setup the Selenium WebDriver (ensure chromedriver is in PATH)
    driver = webdriver.Chrome(options=options)  # You can use other browsers like Firefox or Edge
    driver.get("https://opost.ps/resources/invoices/create")  # Replace with your actual website URL

    try:
        # Wait until the page and input field are loaded
        wait = WebDriverWait(driver, 2)
        input_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@class='multiselect__input']")))

        print("Monitoring the input field. Type shop names...")
        
        while True:
            # Get the current value of the input field
            current_input = input_field.get_attribute("value")
            current_input = current_input.strip().lower()  # Trim spaces and convert to lowercase

            # Check if the input matches any indebted shop
            if current_input and current_input in indebted_shops:
                print(f"ALERT: '{current_input}' is an indebted shop!")
                # Trigger a browser alert
                winsound.Beep(700, 1000)  # Frequency=1000Hz, Duration=500ms
                time.sleep(5)  # Prevent multiple alerts for the same input

            # Optional: Refresh input value periodically
            time.sleep(1)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("End")
        driver.quit()

if __name__ == "__main__":
    monitor_shop_input()