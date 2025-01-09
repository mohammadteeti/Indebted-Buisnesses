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
    print([i for i in indebted_shops],"\n")
    try:
        # Wait for the input field to be loaded
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@class='multiselect__input']")))

        print("Monitoring the input field for copy-paste actions...")

        # Inject JavaScript to detect copy-paste or manual input changes
        driver.execute_script("""
            const inputField = arguments[0];
            inputField.addEventListener('input', function(event) {
                inputField.setAttribute('data-last-input', inputField.value.trim().toLowerCase());
            });
        """, input_field)
        prev_input = "" 
        while True:
            # Check the input field's value
            
            current_input = driver.execute_script("return arguments[0].getAttribute('data-last-input');", input_field)
            print("current_input: " , current_input )
            if prev_input == current_input:
                continue
            if current_input : 
                if current_input.strip() in indebted_shops:
                    print(f"ALERT: '{current_input}' is an indebted shop!")
                    winsound.Beep(700, 1000)  # Trigger sound alert
                    time.sleep(5)  # Prevent duplicate alerts
                    prev_input=current_input
            # Short sleep to prevent high CPU usage
            time.sleep(0.5)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("End")
        winsound.Beep(1000, 500)
        driver.quit()

if __name__ == "__main__":
    monitor_shop_input()