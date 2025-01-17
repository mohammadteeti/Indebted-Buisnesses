import subprocess
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
import psutil
import pygame
debugging_mode_string= ""
url=""
shop_status=""

def start_chrome_session():
    chrome_path = "C:\\Program Files\\Google\\Chrome\\Application"
    os.environ["PATH"] = os.pathsep + chrome_path
    time.sleep(0.5)
    
    # Ensure the debugging command is properly defined
    cmd = debugging_mode_string  # Example: "chrome.exe --remote-debugging-port=9222 --user-data-dir='C:\\selenium\\'"
    
    if not cmd:
        print("Debugging mode command is empty. Check the configuration.")
        return

    try:
        # Use subprocess.Popen to start Chrome in the background
        process = subprocess.Popen(cmd, shell=True)
        print(f"Chrome session started with PID: {process.pid}")
    except Exception as e:
        print(f"Error starting Chrome session: {e}")
    
    
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

    #sheet_number : int= int(input("Enter Sheet Number:"))
    #worksheet = sheet.get_worksheet(sheet_number)  # get worksheet by it's index (index start from 0 for the first sheet)

    sheet_name = input("Enter Sheet Name: ")
    worksheet=sheet.worksheet(sheet_name)

    # Extract data from the 'Business Name' column (Column A)
    shop_names = []

    # Loop over rows, starting from index 1 to second-to-last row in column D
    size=len(worksheet.col_values(4))
    status=worksheet.col_values(4)
    all_shops=worksheet.col_values(1)

    # Extract the names corrosponding to the status = "جاري المتابعة"
    for k in range(1, size- 2):
        if str(status[k]).strip() == str(shop_status):
            shop_names.append(all_shops[k])  # Append the corresponding value from column A




    # Convert all shop names to lowercase and strip spaces
    shop_names = [str(name).strip().lower() for name in shop_names if name]
    return shop_names

# Main function to monitor the list input
def monitor_shop_input():
    # Load the indebted shops from Google Sheet
    SHEET_URL = url #"https://docs.google.com/spreadsheets/d/13EOKkBNrfFFOQBoJF8UxVMCszQVpyntcBL0zNqBWPPc/edit?gid=1956153930#gid=1956153930"
    print("URL:   ", url)
    indebted_shops = load_indebted_shops_from_sheet(SHEET_URL)

    # Connect to an existing Chrome session using the remote debugging port
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"  # Connect to the debugging port
    # Setup the Selenium WebDriver (ensure chromedriver is in PATH)
    driver = webdriver.Chrome(options=options)  # You can use other browsers like Firefox or Edge
    driver.get("https://opost.ps/resources/invoices/create")  # Replace with your actual website URL
    
    #print("\n".join(str(i) for i in indebted_shops))

    print (f"Number of indepted business = {len(indebted_shops)}")

   # with open("indebted_shops.txt","w") as file :
   #     j=1
  #      for i in indebted_shops: 
 #           file.write(f"{j}-{i}\n")
#            j=j+1


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
            #print("current_input: " , current_input )
            if prev_input == current_input:
                continue
            if current_input : 
                if current_input.strip() in indebted_shops:
                    print(f"ALERT: '{current_input}' is an indebted shop!")
                    error_beep.play()
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
        kill_processes_on_port(9222)


def kill_processes_on_port(port):
    """
    Kill all processes using the specified port.
    """
    try:
        for conn in psutil.net_connections(kind="inet"):
            if conn.laddr.port == port:
                pid = conn.pid
                if pid:  # Ensure the connection is associated with a process
                    try:
                        proc = psutil.Process(pid)
                        proc.terminate()  # Graceful termination
                        proc.wait(timeout=3)  # Wait for the process to exit
                        print(f"Terminated process {pid} using port {port}.")
                    except psutil.NoSuchProcess:
                        print(f"Process {pid} does not exist.")
                    except psutil.AccessDenied:
                        print(f"Access denied when trying to terminate process {pid}.")
                    except psutil.TimeoutExpired:
                        proc.kill()  # Force kill if process doesn't terminate
                        print(f"Force-killed process {pid} using port {port}.")
    except Exception as e:
        print(f"An error occurred while killing processes on port {port}: {e}")

if __name__ == "__main__":
    pygame.mixer.init()
    error_beep=  pygame.mixer.Sound("error.wav")
    #start_chrome_session()
    with open("config.cfg", "r",encoding="utf-8") as cfg:
        
        debugging_mode_string = cfg.readline().split(',')[1]
        url = cfg.readline().split(',')[1]
        shop_status = cfg.readline().split(',')[1]
        print(debugging_mode_string)
        print(url)
        print(str(shop_status))
    cfg.close()

    print("starting chrome session ...")
    start_chrome_session()

    print("please login to opost acount !")

    ready = "n"

    while(not(ready.lower()=="y")):
        ready=input ("READY? Y / N ")
        time.sleep(5)

    monitor_shop_input()