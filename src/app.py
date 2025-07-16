import argparse
from datetime import datetime, timedelta
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
                        help="Run for a specific date or range (format: DD-MM-YYYY or DD-MM-YYYY to DD-MM-YYYY)")
    return parser.parse_args()


def get_run_dates(args):
    if args.today:
        return [datetime.today().strftime("%d-%b-%Y")]
    elif args.rundate:
        if "to" in args.rundate:
            start_str, end_str = map(str.strip, args.rundate.split("to"))
            start = datetime.strptime(start_str, "%d-%m-%Y")
            end = datetime.strptime(end_str, "%d-%m-%Y")
            dates = [start + timedelta(days=i) for i in range((end - start).days + 1)]
            return [dt.strftime("%d-%b-%Y") for dt in dates]
        else:
            dt = datetime.strptime(args.rundate, "%d-%m-%Y")
            return [dt.strftime("%d-%b-%Y")]
    else:
        raise ValueError("You must specify either --today or --rundate")


def find_latest_matching_file(folder, prefix):
    pattern = os.path.join(folder, f"{prefix}*.xlsx")
    matches = glob.glob(pattern)
    if not matches:
        return None
    matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return matches[0]


def run_modes(modes, run_dates, download=False):
    config = load_config()
    RM_SHEET_CONFIG = config.get("RM_SHEET_CONFIG", {})
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    skipped_downloads = set()
    charge_file = None

    # If any download is needed
    if download or "charge" in modes:
        print("\n🌐 Launching browser for DSM file download...")
        driver = setup_browser_driver()
        wait = WebDriverWait(driver, config.get("default_timeout", 180))

        # Login to DSM
        login_dsm(driver, wait, config["dsm"]["url"], config["dsm"]["user"], config["dsm"]["password"])

        # Only use first run_date for charge
        rundate_obj = datetime.strptime(run_dates[0], "%d-%b-%Y")
        charge_filename = f"CHARGE_AND_DUMP_REPORT_{rundate_obj.day}_{rundate_obj.month}_{rundate_obj.year}.xlsx"

        # Call download logic (hourly section only downloads this file)
        skipped_downloads = go_to_file_station_and_download(
            driver, wait,
            config["download_filenames"],
            config["dsm"]["file_station"],
            config["dsm"]["hourly_url"],
            selected_modes=modes,
            run_date=run_dates[0],
            target_filename=charge_filename
        ) or set()

        driver.quit()

    # For RM and DPR, look in downloaded folder
    download_folder = config["download_folder"]
    dpr_file = None if "BF-02 DPR" in skipped_downloads else find_latest_matching_file(download_folder, "BF-02 DPR")
    rm_file = None if "11A BF-02 BUNKER" in skipped_downloads else find_latest_matching_file(download_folder, "11A BF-02 BUNKER")
    charge_file = None if "charge_and_dump" in skipped_downloads else find_latest_matching_file(download_folder, "CHARGE_AND_DUMP_REPORT_")

    for run_date in run_dates:
        # Process RM
        if "rm" in modes:
            if not rm_file:
                print("❌ Bunker (RM) file not found in download folder or skipped.")
            else:
                print(f"📦 Processing RM sheet for {run_date}")
                read_rm_sheet(
                    file_path=rm_file,
                    RM_SHEET_CONFIG=RM_SHEET_CONFIG,
                    start_date=run_date,
                    output_dir=output_dir
                )

        # Process DPR
        if "dpr" in modes:
            if not dpr_file:
                print("❌ DPR file not found in download folder or skipped.")
            else:
                print(f"📊 Processing DPR sheet for {run_date}")
                update_dpr_config_from_excel(
                    dpr_file,
                    os.path.join("src", "config", "setting.yaml"),
                    run_date
                )
                read_dpr_sheet(
                    file_path=dpr_file,
                    config=config,
                    start_date=run_date,
                    output_dir=output_dir
                )

    # Merge charge report
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

        run_dates = get_run_dates(args)
        run_modes(modes, run_dates, download=args.today)
    except Exception as e:
        print(f"❌ Error: {e}")