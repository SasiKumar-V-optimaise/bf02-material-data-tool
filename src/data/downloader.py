import os
import sys
import time
import pandas as pd
from dotenv import load_dotenv
from selenium.webdriver.support.ui import WebDriverWait
from config.loader import load_config
from src.utils.helper_functions_downloader import (
    setup_browser_driver,
    login_dsm,
    go_to_file_station_and_download,
    read_sheets_by_config,
)
from datetime import datetime

config = load_config()
sheet_configs = config.get("sheet_configs", {})

# Load credentials
load_dotenv()
LOGIN_URL = os.getenv("EVONITH_URL")
FILESTATION_URL = os.getenv("EVONITH_FILESTATION_URL")
USER = os.getenv("EVONITH_USER")
PASSWD = os.getenv("EVONITH_PASS")
DEFAULT_TIMEOUT = 180

# Main script
if __name__ == "__main__":
    # Main execution block: sets up the browser, logs in, downloads the file, processes it, and closes the browser.
    driver = setup_browser_driver()
    wait = WebDriverWait(driver, DEFAULT_TIMEOUT)
    try:
        login_dsm(driver, wait)
        go_to_file_station_and_download(driver, wait)
        filename = config.get("download_filename", "11A BF-02 BUNKER 2025-26.xlsx")
        download_folder = config.get("download_folder", os.path.join(os.getcwd(), "downloads"))
        downloaded_file = os.path.join(download_folder, filename) 
        # TODO: Make this file path configurable
        if os.path.exists(downloaded_file):
            read_sheets_by_config(downloaded_file, sheet_configs, start_date="24-Jun-2025")
        else:
            print(f"⚠️ File not found: {downloaded_file}")

        input("📁 Done. Press Enter to close browser...")
    finally:
        driver.quit()



