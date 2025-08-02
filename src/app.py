import argparse
from datetime import datetime, timedelta
import os
import glob
import pandas as pd
from config.loader import load_config
from utils.influx_writer import push_dataframe_to_influx
from utils.helper_functions_downloader import (
    read_rm_sheet,
    read_dpr_sheet,
    setup_browser_driver,
    login_eml,
    go_to_file_station_and_download,
    update_dpr_config_from_excel,
    process_shiftwise_charge_data
)
from selenium.webdriver.support.ui import WebDriverWait

def rename_fields(df, field_mapping):
    """
    Rename dataframe columns using the field_mapping defined in YAML config.
    Ignores fields not present in mapping.
    """
    return df.rename(columns={k: v for k, v in field_mapping.items() if k in df.columns})

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
    print("🔍 Config keys loaded:", config.keys())

    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)

    skipped_downloads = set()
    charge_file = None

    # If any download is needed
    if download or "charge" in modes:
        print("\n🌐 Launching browser for EML file download...")
        driver = setup_browser_driver()
        wait = WebDriverWait(driver, config.get("default_timeout", 180))

        # Login to DSM
        login_eml(driver, wait, config["eml"]["url"], config["eml"]["user"], config["eml"]["password"])

        # Only use first run_date for charge
        rundate_obj = datetime.strptime(run_dates[0], "%d-%b-%Y")
        charge_filename = f"CHARGE_AND_DUMP_REPORT_{rundate_obj.day}_{rundate_obj.month}_{rundate_obj.year}.xlsx"

        # Call download logic (hourly section only downloads this file)
        skipped_downloads = go_to_file_station_and_download(
            driver, wait,
            config["download_filenames"],
            config["eml"]["file_station"],
            config["eml"]["hourly_url"],
            selected_modes=modes,
            run_date=run_dates[0],
            target_filename=charge_filename
        ) or set()

        driver.quit()

    # For RM and DPR, look in downloaded folder
    download_folder = config["download_folder"]
    dpr_file = None if "BF-02 DPR" in skipped_downloads else find_latest_matching_file(download_folder, "BF-02 DPR")
    rm_file = None if "11A BF-02 BUNKER" in skipped_downloads else find_latest_matching_file(download_folder, "11A BF-02 BUNKER")

    def find_charge_files_for_shift(run_date_str, download_folder):
        run_date = datetime.strptime(run_date_str, "%d-%b-%Y")
        prev_date = run_date - timedelta(days=1)

        def find_latest_file_for_date(date_obj):
            pattern = os.path.join(
                download_folder,
                f"CHARGE_AND_DUMP_REPORT_{date_obj.day}_{date_obj.month}_{date_obj.year}*.xlsx"
            )
            matches = glob.glob(pattern)
            if not matches:
                return None
            matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return matches[0]

        current_file = find_latest_file_for_date(run_date)
        previous_file = find_latest_file_for_date(prev_date)

        return current_file, previous_file
    
    # Function to rename fields in a DataFrame
    def rename_fields(df: pd.DataFrame, mapping: dict):
        return df.rename(columns=mapping)
    # Process RM
    if "rm" in modes:
        if not rm_file:
            print("❌ Bunker (RM) file not found in download folder or skipped.")
        else:
            print(f"📦 Processing RM sheet for {run_dates[0]}" if len(run_dates) == 1 else f"📦 Processing RM sheet from {run_dates[0]} to {run_dates[-1]}")
            
            read_rm_sheet(
                file_path=rm_file,
                RM_SHEET_CONFIG=RM_SHEET_CONFIG,
                start_date=run_dates,
                output_dir=output_dir
            )

            # Write RM to InfluxDB
            influx_cfg = config["influxdb"]
            rm_df_path = os.path.join(output_dir, "combined_bunker_data.xlsx")
            if os.path.exists(rm_df_path):
                df_rm = pd.read_excel(rm_df_path)

                # Apply field mapping
                field_mapping = config.get("rm_feilds", {})  # Make sure spelling is rm_feilds as in your YAML
                df_rm = rename_fields(df_rm, field_mapping)
                # df_rm = clean_and_convert(df_rm, list(config["rm_feilds"].keys()))

                # Ensure 'date' column exists
                if "date" not in df_rm.columns:
                    raise ValueError("Missing 'date' column after renaming RM fields.")
                
                # Push to InfluxDB
                push_dataframe_to_influx( df_rm, influx_cfg["bucket"], "rm_data", influx_cfg, field_mapping=config["rm_feilds"])
                print("✅ RM data pushed to InfluxDB.")

    # 📊 Process DPR for all run_dates
    if "dpr" in modes:
        if not dpr_file:
            print("❌ DPR file not found in download folder or skipped.")
        else:
            combined_dpr_dfs = []
            config_cache = {}  # cache config for (month, year)

            for run_date in run_dates:
                print(f"📊 Processing DPR sheet for {run_date}")
                run_date_obj = datetime.strptime(run_date, "%d-%b-%Y")
                month_year_key = (run_date_obj.month, run_date_obj.year)

                if month_year_key not in config_cache:
                    print(f"🔁 Updating config for {run_date}")
                    update_dpr_config_from_excel(
                        dpr_file,
                        os.path.join("src", "config", "setting.yaml"),
                        run_date
                    )

                config = load_config(os.path.join("src", "config", "setting.yaml"))
                config_cache[month_year_key] = config  # save loaded config for reuse

                # Read the data
                df = read_dpr_sheet(
                    file_path=dpr_file,
                    config=config,
                    start_date=run_date,
                    output_dir=output_dir
                )

                if df is not None:
                    combined_dpr_dfs.append(df)

            if combined_dpr_dfs:
                final_df = pd.concat(combined_dpr_dfs, ignore_index=True)
                out_path = os.path.join(output_dir, "combined_dpr_data.xlsx")
                final_df.to_excel(out_path, index=False)
                print(f"\n✅ Final DPR data written → {out_path}")

                # Use the last config used (all are same for same month/year)
                field_mapping = config.get("dpr_fields", {})
                final_df = rename_fields(final_df, field_mapping)

                if "date" not in final_df.columns:
                    raise ValueError("Missing 'date' column after renaming DPR fields.")

                push_dataframe_to_influx(final_df, config["influxdb"]["bucket"], "dpr_data", config["influxdb"], field_mapping=config["dpr_fields"])
                print("✅ DPR data pushed to InfluxDB.")
            else:
                print("⚠️ No DPR data found for any of the dates.")

    # Merge charge report
    if "charge" in modes:
        charge_file_current, charge_file_prev = find_charge_files_for_shift(run_dates[0], download_folder)
        if not charge_file_current:
            print("❌ Charge report file not found.")
        else:
            print(f"🔁 Processing charge reports: {charge_file_prev} and {charge_file_current}")
            process_shiftwise_charge_data(charge_file_current, charge_file_prev, output_dir, run_dates[0])

            

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