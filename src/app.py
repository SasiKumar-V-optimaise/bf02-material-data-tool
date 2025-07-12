import argparse
from datetime import datetime
import os
import glob
from config.loader import load_config
from utils.helper_functions_downloader import (
    read_rm_sheet,
    read_dpr_sheet,
    merge_dpr_and_bunker,
    merge_hourly_excel,
    setup_browser_driver,
    login_dsm,
    go_to_file_station_and_download,
    update_dpr_config_from_excel
)
from selenium.webdriver.support.ui import WebDriverWait


def parse_args():
    parser = argparse.ArgumentParser(description="Offline Data Consolidation CLI")
    parser.add_argument("--mode", required=True,
                        help="Mode(s) to run: rm, dpr, charge, rm,dpr, or both")
    parser.add_argument("--today", action="store_true",
                        help="Use today’s date as run date")
    parser.add_argument("--rundate", type=str,
                        help="Run for a specific date (format: DD-MM-YYYY)")
    return parser.parse_args()


def get_run_date(args):
    if args.today:
        return datetime.today().strftime("%d-%b-%Y")
    elif args.rundate:
        dt = datetime.strptime(args.rundate, "%d-%m-%Y")
        return dt.strftime("%d-%b-%Y")
    else:
        raise ValueError("You must specify either --today or --rundate")


def find_latest_matching_file(folder, prefix):
    pattern = os.path.join(folder, f"{prefix}*.xlsx")
    matches = glob.glob(pattern)
    if not matches:
        return None
    matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return matches[0]


def run_modes(modes, run_date, download=False):
    config = load_config()
    RM_SHEET_CONFIG = config.get("RM_SHEET_CONFIG", {})
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    # Track skipped downloads
    skipped_downloads = set()

    if download:
        print("\n🌐 Logging into DSM and downloading files...")
        driver = setup_browser_driver()
        wait = WebDriverWait(driver, config.get("default_timeout", 180))
        login_dsm(driver, wait, config["dsm"]["url"], config["dsm"]["user"], config["dsm"]["password"])
        skipped_downloads = go_to_file_station_and_download(
            driver, wait, config["download_filenames"],
            config["dsm"]["file_station"], config["dsm"]["hourly_url"],
            selected_modes=modes
        ) or set()
        driver.quit()

    download_folder = config["download_folder"]
    dpr_file = None if "BF-02 DPR" in skipped_downloads else find_latest_matching_file(download_folder, "BF-02 DPR")
    rm_file = None if "11A BF-02 BUNKER" in skipped_downloads else find_latest_matching_file(download_folder, "11A BF-02 BUNKER")
    charge_file = None if "charge_and_dump" in skipped_downloads else find_latest_matching_file(download_folder, "charge_and_dump")

    if "rm" in modes:
        if not rm_file:
            print("❌ Bunker (RM) file not found in download folder or skipped.")
        else:
            print(f"📦 Processing RM sheet: {rm_file}")
            read_rm_sheet(
                file_path=rm_file,
                RM_SHEET_CONFIG=RM_SHEET_CONFIG,
                start_date=run_date,
                output_dir=output_dir
            )

    if "dpr" in modes:
        if not dpr_file:
            print("❌ DPR file not found in download folder or skipped.")
        else:
            print(f"📊 Processing DPR sheet: {dpr_file}")
            update_dpr_config_from_excel(
                dpr_file,
                os.path.join("src", "config", "setting.yaml"),
                run_date  # ✅ pass the run date here
            )
            read_dpr_sheet(
                file_path=dpr_file,
                config=config,
                start_date=run_date,
                output_dir=output_dir
            )

    if "rm" in modes and "dpr" in modes:
        dpr_path = os.path.join(output_dir, "combined_dpr_data.xlsx")
        bunker_path = os.path.join(output_dir, "combined_bunker_data.xlsx")
        yaml_path = os.path.join("src", "config", "setting.yaml")
        final_output = "final_combined_data.xlsx"

        if os.path.exists(dpr_path) and os.path.exists(bunker_path):
            print(f"🔗 Merging: {os.path.basename(dpr_path)} + {os.path.basename(bunker_path)}")
            merge_dpr_and_bunker(
                dpr_path=dpr_path,
                bunker_path=bunker_path,
                yaml_path=yaml_path,
                master_path=final_output
            )
        else:
            print("⚠️ Cannot merge — one or both inputs missing.")

    if "charge" in modes:
        if not charge_file:
            print("❌ Charge report not found or skipped.")
        else:
            print(f"🔁 Merging Charge Report: {charge_file}")
            merge_hourly_excel(charge_file)


if __name__ == "__main__":
    args = parse_args()
    try:
        modes_raw = args.mode.lower()
        if "both" in modes_raw:
            modes = ["rm", "dpr"]
        else:
            modes = [m.strip() for m in modes_raw.split(",")]

        run_date = get_run_date(args)
        run_modes(modes, run_date, download=args.today)
    except Exception as e:
        print(f"❌ Error: {e}")
