import os
import time
from selenium.webdriver.support.ui import WebDriverWait
from config.loader import load_config
from utils.helper_functions_downloader import (
    setup_browser_driver,
    login_dsm,
    go_to_file_station_and_download,
    read_rm_sheet,
    read_dpr_sheet,
    merge_dpr_and_bunker,
    merge_hourly_excel,
)
from datetime import datetime

# Load YAML config
config          = load_config()
files_to_get    = config["download_filenames"]
download_folder = config["download_folder"]
RM_SHEET_CONFIG = config["RM_SHEET_CONFIG"]

# DSM credentials & URLs
dsm_cfg         = config["dsm"]
LOGIN_URL       = dsm_cfg["url"]
FILESTATION_URL = dsm_cfg["file_station"]
USER            = dsm_cfg["user"]
PASSWD          = dsm_cfg["password"]

DEFAULT_TIMEOUT = config.get("default_timeout", 180)
START_DATE      = config.get("start_date", "01-Jun-2025")

if __name__ == "__main__":
    driver = setup_browser_driver()
    wait   = WebDriverWait(driver, DEFAULT_TIMEOUT)

    try:
        login_dsm(driver, wait,LOGIN_URL, USER, PASSWD)

        # ⬇️ Download files & get name of latest HOURLY .xlsx file
        latest_hourly_file = go_to_file_station_and_download(driver, wait, files_to_get)


        # # ⬇️ Process DPR and RM sheets
        # for fname in files_to_get:
        #     path = os.path.join(download_folder, fname)
        #     if not os.path.exists(path):
        #         print(f"⚠️ File not found: {path} — skipping")
        #         continue

        #     if "DPR" in fname.upper():
        #         print(f"📄 Processing DPR sheet: {fname}")
        #         read_dpr_sheet(path, config=config, output_dir="outputs")
        #     else:
        #         print(f"📄 Processing RM sheet: {fname}")
        #         read_rm_sheet(
        #             file_path=path,
        #             RM_SHEET_CONFIG=RM_SHEET_CONFIG,
        #             start_date=START_DATE,
        #             output_dir="outputs",
        #         )

        # ✅ Merge DPR and Bunker outputs if needed
        # try:
        #     dpr_path = os.path.join("outputs", "combined_dpr_Jun25.xlsx")
        #     bunker_path = os.path.join("outputs", "combined_bunker_data.xlsx")
        #     final_output = "final_combined_data.xlsx"

        #     if os.path.exists(dpr_path) and os.path.exists(bunker_path):
        #         merge_dpr_and_bunker(dpr_path, bunker_path, final_output)
        #     else:
        #         print("⚠️ One or both files not found for merging.")
        # except Exception as e:
        #     print(f"❌ Error during merging: {e}")
        
        # ⬇️ Merge hourly Excel files        # ⬇️ Merge hourly file if available
        if latest_hourly_file:
            hourly_path = os.path.join(download_folder, latest_hourly_file)

            # Optional: Wait for file to be fully downloaded
            timeout = 30
            while not os.path.exists(hourly_path) and timeout > 0:
                time.sleep(1)
                timeout -= 1

            if os.path.exists(hourly_path):
                print(f"📥 Merging latest hourly Excel file: {hourly_path}")
                merge_hourly_excel(hourly_path)
            else:
                print(f"❌ Downloaded file not found in: {hourly_path}")
        else:
            print("⚠️ No latest hourly .xlsx file found.")


    finally:
        driver.quit()
