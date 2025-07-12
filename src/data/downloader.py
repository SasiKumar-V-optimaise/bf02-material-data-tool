import os
import time
import glob
from datetime import datetime
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
    update_dpr_config_from_excel,
    load_metadata,

)

# === CONFIGURATION ===
config            = load_config()
files_to_get      = config["download_filenames"]
download_folder   = config["download_folder"]
RM_SHEET_CONFIG   = config["RM_SHEET_CONFIG"]
ROOT_URL          = config["dsm"]["file_station"]
HOURLY_URL        = config["dsm"]["hourly_url"]
DEFAULT_TIMEOUT   = config.get("default_timeout", 180)
START_DATE        = config.get("start_date", "01-Jun-2025")

dsm_cfg           = config["dsm"]
LOGIN_URL         = dsm_cfg["url"]
FILESTATION_URL   = dsm_cfg["file_station"]
USER              = dsm_cfg["user"]
PASSWD            = dsm_cfg["password"]

# === UTILITY: Find latest matching file ===
def find_latest_matching_file(folder, prefix):
    """
    Finds the latest modified .xlsx file in the folder that starts with the given prefix.
    """
    pattern = os.path.join(folder, f"{prefix}*.xlsx")
    matches = glob.glob(pattern)
    if not matches:
        return None
    matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return matches[0]


# === MAIN EXECUTION ===
if __name__ == "__main__":
    driver = setup_browser_driver()
    wait = WebDriverWait(driver, DEFAULT_TIMEOUT)

    try:
        login_dsm(driver, wait, LOGIN_URL, USER, PASSWD)

        # ⬇️ Download files from Evonith Shared Folder
        latest_hourly_file = go_to_file_station_and_download(driver, wait, files_to_get, ROOT_URL, HOURLY_URL)

        # ⬇️ Process downloaded DPR or RM files
        for fname_prefix in files_to_get:
            latest_file = find_latest_matching_file(download_folder, fname_prefix)

            if not latest_file:
                print(f"⚠️ No file found for prefix: {fname_prefix} — skipping.")
                continue

            if "DPR" in fname_prefix.upper():
                # 🆕 Update YAML config with the newest sheet before parsing
                yaml_file = os.path.join("src", "config", "setting.yaml")
                update_dpr_config_from_excel(latest_file, yaml_file)

                print(f"📄 Processing DPR sheet: {os.path.basename(latest_file)}")
                read_dpr_sheet(latest_file, config=config, output_dir="outputs")
            else:
                print(f"📄 Processing RM sheet: {os.path.basename(latest_file)}")
                read_rm_sheet(
                    file_path=latest_file,
                    RM_SHEET_CONFIG=RM_SHEET_CONFIG,
                    start_date=START_DATE,
                    output_dir="outputs",
                )


        # Merge DPR + Bunker
        dpr_files = glob.glob(os.path.join("outputs", "combined_dpr_*.xlsx"))
        bunker_path = os.path.join("outputs", "combined_bunker_data.xlsx")
        yaml_path = os.path.join("src", "config", "setting.yaml")
        final_output = "final_combined_data.xlsx"

        if dpr_files and os.path.exists(bunker_path):
            latest_dpr_file = max(dpr_files, key=os.path.getmtime)  # pick most recent
            merge_dpr_and_bunker(
                dpr_path=latest_dpr_file,
                bunker_path=bunker_path,
                yaml_path=yaml_path,
                master_path=final_output
            )
        else:
            print("⚠️ Required files not found for merging.")

        # ⬇️ Merge latest HOURLY Excel if downloaded
        # Fallback: Use filename from metadata if not downloaded freshly
        previous_metadata = load_metadata()
        latest_hourly_files = previous_metadata.get("HOURLY_REPORT", {}).get("name")

        # ⬇️ Merge latest HOURLY Excel if downloaded or available from metadata
        # if latest_hourly_files:
        #     hourly_path = os.path.join(download_folder, latest_hourly_file)

        #     # Wait for file to finish downloading (also wait if .crdownload exists)
        #     timeout = 30
        #     while (not os.path.exists(hourly_path) or
        #         glob.glob(os.path.join(download_folder, "*.crdownload"))) and timeout > 0:
        #         time.sleep(1)
        #         timeout -= 1

        #     if os.path.exists(hourly_path):
        #         print(f"📥 Merging latest hourly Excel file: {hourly_path}")
        #         merge_hourly_excel(hourly_path)
        #     else:
        #         print(f"❌ Downloaded HOURLY file not found: {hourly_path}")
        # else:
        #     print("⚠️ No latest HOURLY .xlsx file found.")


    finally:
        driver.quit()
