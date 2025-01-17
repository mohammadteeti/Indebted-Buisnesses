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
import os

debugging_mode_string = ""
url = ""
shop_status = ""

def start_edge_session():
    cmd = debugging_mode_string
    os.system(cmd)

# Google Sheets setup
def load_indebted_shops_from_sheet(sheet_url):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_url(sheet_url)

    sheet_name = input("Enter Sheet Name: ")
    worksheet = sheet.worksheet(sheet_name)

    shop_names = []
    size = len(worksheet.col_values(4))
    status = worksheet.col_values(4)
    all_shops = worksheet.col_values(1)

    for k in range(1, size - 2):
        if str(status[k]).strip() == str(shop_status) + " ":
            shop_names.append(all_shops[k])

    shop_names = [str(name).strip().lower() for name in shop_names if name]
    return shop_names

# Main function to monitor the list input
def monitor_shop_input():
    SHEET_URL = url
    print("URL:   ", url)
    indebted_shops = load_indebted_shops_from_sheet(SHEET_URL)

    options = webdriver.EdgeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # Connect to the debugging port

    driver = webdriver.Edge(options=options)  # Use Edge WebDriver
    driver.get("https://opost.ps/resources/invoices/create")

    print(f"Number of indebted businesses = {len(indebted_shops)}")

    try:
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@class='multiselect__input']")))

        print("Monitoring the input field for copy-paste actions...")

        driver.execute_script("""
            const inputField = arguments[0];
            inputField.addEventListener('input', function(event) {
                inputField.setAttribute('data-last-input', inputField.value.trim().toLowerCase());
            });
        """, input_field)

        prev_input = ""
        while True:
            current_input = driver.execute_script("return arguments[0].getAttribute('data-last-input');", input_field)
            if prev_input == current_input:
                continue
            if current_input:
                if current_input.strip() in indebted_shops:
                    print(f"ALERT: '{current_input}' is an indebted shop!")
                    winsound.Beep(700, 1000)
                    time.sleep(5)
                    prev_input = current_input
            time.sleep(0.5)

    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("End")
        winsound.Beep(1000, 400)
        time.sleep(0.5)
        winsound.Beep(1000, 400)
        driver.quit()

if __name__ == "__main__":
    with open("config_Edge_version.cfg", "r") as cfg:
        
        debugging_mode_string = cfg.readline().split(',')[1]
        url = cfg.readline().split(',')[1]
        shop_status = cfg.readline().split(',')[1]
        print("debugging_mode_string",debugging_mode_string)
        print(url)
        print(str(shop_status))
    cfg.close()

    print("starting Edge session ...")
    start_edge_session()

    print("please login to opost account!")

    ready = "n"

    while not ready.lower() == "y":
        ready = input("READY? Y / N ")
        time.sleep(1)

    monitor_shop_input()

