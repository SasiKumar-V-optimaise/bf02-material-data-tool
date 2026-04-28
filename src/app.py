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
    ChargeDataProcessor,
    update_hot_metal_config_from_excel,
    read_hot_metal_sheet,
    process_rm_hm_sheet,
)
from selenium.webdriver.support.ui import WebDriverWait
import logging

# Configure logging (console + timestamp)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def rename_fields(df, field_mapping):
    """
    Rename dataframe columns using the field_mapping defined in YAML config.
    Ignores fields not present in mapping.
    """
    return df.rename(columns={k: v for k, v in field_mapping.items() if k in df.columns})

def parse_args():
    parser = argparse.ArgumentParser(description="Offline Data Consolidation CLI")
    parser.add_argument("--mode", required=True,
                        help="Mode(s) to run: rm, dpr, charge, hot_metal, or both")
    parser.add_argument("--today", action="store_true",
                        help="Use todayâs date as run date")
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
        logger.info("Launching browser for EML file download...")
        driver = setup_browser_driver()
        wait = WebDriverWait(driver, config.get("default_timeout", 180))

        # Login to DSM
        login_eml(driver, wait, config["eml"]["url"], config["eml"]["user"], config["eml"]["password"])

        for run_date_str in run_dates:
            rundate_obj = datetime.strptime(run_date_str, "%d-%b-%Y")
            charge_filename = f"CHARGE_AND_DUMP_REPORT_{rundate_obj.day}_{rundate_obj.month}_{rundate_obj.year}.xlsx"
            # logger.info("Downloading charge report for %s: %s", run_date_str, charge_filename)

            # Download charge report for this date
            skipped_downloads = go_to_file_station_and_download(
                driver, wait,
                config["download_filenames"],
                config["eml"]["file_station"],
                config["eml"]["hourly_url"],
                selected_modes=modes,
                run_date=run_date_str,
                target_filename=charge_filename
            ) or set()

        driver.quit()


    # For RM and DPR, look in downloaded folder
    download_folder = config["download_folder"]
    dpr_file = None if "BF-02 DPR" in skipped_downloads else find_latest_matching_file(download_folder, "BF-02 DPR")
    rm_file = None if "01E BF-02 BUNKER  2026-27" in skipped_downloads else find_latest_matching_file(download_folder, "01E BF-02 BUNKER  2026-27")

    def find_charge_files(run_date_str, download_folder):
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
            logger.error("Bunker (RM) file not found in download folder or skipped.")
        else:
            for run_date in run_dates:
                logger.info(f"Processing RM sheet for {run_date}")

                # Read RM sheet for this date
                read_rm_sheet(
                    file_path=rm_file,
                    RM_SHEET_CONFIG=RM_SHEET_CONFIG,
                    start_date=[run_date],  # pass single date in list
                    output_dir=output_dir
                )

                # Path to combined Excel
                rm_df_path = os.path.join(output_dir, "combined_bunker_data.xlsx")
                if os.path.exists(rm_df_path):
                    df_rm = pd.read_excel(rm_df_path)

                    # Apply field mapping
                    field_mapping = config.get("rm_feilds", {})  # spelling as in YAML
                    df_rm = rename_fields(df_rm, field_mapping)

                    # Ensure 'date' column exists
                    if "date" not in df_rm.columns:
                        logger.error(f"Missing 'date' column after renaming RM fields for {run_date}. Skipping this date.")
                        continue

                    # Push to InfluxDB
                    influx_cfg = config["influxdb"]
                    try:
                        push_dataframe_to_influx(
                            df=df_rm,
                            bucket=influx_cfg["bucket"],
                            measurement="rm_updated_data",
                            influx_config=influx_cfg,
                            field_mapping=field_mapping
                        )
                        logger.info(f"RM data pushed to InfluxDB for {run_date}.")
                    except Exception as e:
                        logger.error(f" Failed to push RM data for {run_date}: {e}")
                else:
                    logger.warning(f"Combined RM file not found for {run_date}, skipping.")

   
    #  Process DPR for all run_dates
    if "dpr" in modes:
        if not dpr_file:
            logger.warning(" DPR file not found in download folder or skipped.")
        else:
            config_cache = {}  # cache config for (month, year)
            for run_date in run_dates:
                logger.info(f"Processing DPR sheet for {run_date}")
                run_date_obj = datetime.strptime(run_date, "%d-%b-%Y")
                month_year_key = (run_date_obj.month, run_date_obj.year)

                if month_year_key not in config_cache:
                    logger.info(f"Updating config for {run_date}")
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

                if df is not None and not df.empty:
                    field_mapping = config.get("dpr_fields", {})
                    df = rename_fields(df, field_mapping)
                    # Save daily Excel (optional)
                    out_path = os.path.join(output_dir, f"combined_dpr_data.xlsx")
                    df.to_excel(out_path, index=False)
                    logger.info(f"DPR data written â {out_path}")

                    # Rename columns
                    

                    if "date" not in df.columns:
                        raise ValueError("Missing 'date' column after renaming DPR fields.")

                    # Push this day's data directly to Influx
                    push_dataframe_to_influx(
                        df,
                        config["influxdb"]["bucket"],
                        "dpr_data",
                        config["influxdb"],
                        field_mapping=config["dpr_fields"]
                    )
                    logger.info(f" DPR data for {run_date} pushed to InfluxDB.\n")
                else:
                    logger.warning(f" No DPR data found for {run_date}.")

    
    # Process Hot Metal
    if "hot_metal" in modes:
        base_cfg, hm_cfg = os.path.join("src", "config", "setting.yaml"), os.path.join("src", "config", "hot_metal.yaml")
        hm_file = None if "06 BF-02- HOT METAL, SLAG & GAS" in skipped_downloads else find_latest_matching_file(download_folder, "01B BF-02- HOT METAL, SLAG & GAS")

        if not hm_file:
            logger.info(" HOT_METAL file not found in download folder or skipped.")
        else:
            logger.info(f" Using HOT_METAL file â {os.path.basename(hm_file)}")
            fmap = None
            for rundate in run_dates:
                try:
                    # update config for that date
                    update_hot_metal_config_from_excel(hm_file, hm_cfg, rundate)
                    cfg_hm = load_config(hm_cfg)

                    # read single-date data
                    df = read_hot_metal_sheet(hm_file, [rundate], cfg_hm, output_dir=None)
                    if df is None or df.empty:
                        logger.warning(f" No data for {rundate}, skipping...")
                        continue

                    # cleanup
                    fmap = cfg_hm.get("hot_metal_fields", {})
                    if "DATE" in df.columns and "date" in df.columns:
                        df = df.drop(columns=["DATE"])
                    df = rename_fields(df, fmap).loc[:, ~pd.Index.duplicated(df.columns)]

                    if "date" not in df.columns:
                        raise ValueError("Missing 'date' after renaming.")
                    if df is not None and not df.empty:
                        field_mapping = config.get("dpr_fields", {})
                        df = rename_fields(df, field_mapping)
                        # Save daily Excel (optional)
                        out_path = os.path.join(output_dir, f"combined_hot_data.xlsx")
                        df.to_excel(out_path, index=False)
                        logger.info(f"DPR data written â {out_path}")
                    
                    # convert some columns to tag strings
                    for col in ["lab_sample_id", "cast_no_ladle_spec"]:
                        if col in df.columns:
                            df[col] = df[col].astype(str).fillna("")

                    # write directly to Influx (daily push like DPR)
                    push_dataframe_to_influx(
                        df,
                        load_config(base_cfg)["influxdb"]["bucket"], 
                        "hotmetal_slag_updated_data",
                        load_config(base_cfg)["influxdb"],
                        field_mapping=fmap,
                        tag_keys=["lab_sample_id", "cast_no_ladle_spec"]
                    )
                    logger.info(f" HOT_METAL {rundate} written to InfluxDB.")

                except Exception as e:
                    logger.info(f" HOT_METAL {rundate}: {e}")

        # ---------------------------------------------------------------------
    # Process RM & HM Combined File
    # ---------------------------------------------------------------------
    if "rm_hm" in modes:
        logger.info("Starting RM & HM data processing...")

        # Locate the RM & HM file
        rm_hm_file = None if "RM & HM" in skipped_downloads else find_latest_matching_file(download_folder, "RM & HM")

        if not rm_hm_file:
            logger.warning("RM & HM file not found in download folder or skipped.")
        else:
            for run_date in run_dates:
                try:
                    logger.info(f"Processing RM & HM data for {run_date}")

                    # Read RM & HM sheet (new helper function)
                    df_rmhm = process_rm_hm_sheet(
                        file_path=rm_hm_file,
                        config=config,
                        start_date=[run_date],
                        output_dir=output_dir
                    )

                    if df_rmhm is None or df_rmhm.empty:
                        logger.warning(f"No RM & HM data found for {run_date}")
                        continue

                    # Rename fields based on YAML config
                    field_mapping = config.get("rm_hm_fields", {})
                    df_rmhm = rename_fields(df_rmhm, field_mapping)

                    # Ensure date column
                    if "date" not in df_rmhm.columns:
                        raise ValueError("Missing 'date' column after renaming RM & HM fields.")

                    # Push to InfluxDB
                    influx_cfg = config["influxdb"]
                    push_dataframe_to_influx(
                        df_rmhm,
                        bucket=influx_cfg["bucket"],
                        measurement="rm_hm_data",
                        influx_config=influx_cfg,
                        field_mapping=field_mapping
                    )

                    logger.info(f"RM & HM data for {run_date} pushed to InfluxDB.")

                except Exception as e:
                    logger.error(f"Error processing RM & HM for {run_date}: {e}")



    # Merge charge report

    if "charge" in modes:
        config = load_config(os.path.join("src", "config", "setting.yaml"))


        # Neon DB config
        neon_cfg = {
            "host": "ep-silent-bush-abgf2mw2-pooler.eu-west-2.aws.neon.tech",
            "dbname": "neondb",
            "user": "neondb_owner",
            "password": "npg_o2m8qDpOlaAF",
            "sslmode": "require",
        }

        for run_date_str in run_dates:
            charge_file_current, charge_file_prev = find_charge_files(run_date_str, download_folder)
            if not charge_file_current:
                logger.error("Charge report file not found for %s.", run_date_str)
                continue

            logger.info("Processing charge reports for %s: %s and %s",
                        run_date_str,
                        charge_file_prev if charge_file_prev else "None",
                        charge_file_current)

            processor = ChargeDataProcessor(
                file_today=charge_file_current,
                file_yesterday=charge_file_prev,
                output_dir=output_dir,
                run_date_str=run_date_str,
                neon_cfg=neon_cfg,     
                influx_cfg={
                    "url": "https://eu-central-1-1.aws.cloud2.influxdata.com",
                    "token": "yZNDCGAqOrCP4HFdzDashFQfNhqNJRIxB6Q4atvNUoV8Zt2jEsO-eS-T57U2crSsp-GMv9HMBwoVELA6aTM_lQ==",
                    "org": "Blast Furnace, Evonith"
                },
                material_groups=config.get("material_groups",{})
            )

            df = processor.process()

            if df is not None and not df.empty:
                df = df.rename(columns={"DATETIME": "date"})

                if "date" not in df.columns:
                    raise ValueError("Missing 'date' column after processing charge data.")

                

                field_mapping = config.get("charge_fields", {})


                if field_mapping:
                    df = rename_fields(df, field_mapping)

                df.to_excel(os.path.join(output_dir, f"charge_data_{run_date_str}.xlsx"), index=False)
                push_dataframe_to_influx(
                    df,
                    config["influxdb"]["bucket"],
                    "latest_charge_data",
                    config["influxdb"],
                )

                logger.info("Charge data for %s pushed to InfluxDB.", run_date_str)

            else:
                logger.warning("No charge data found for %s.", run_date_str)

                    
import traceback
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
        logger.warning("Exception occurred:\n" + traceback.format_exc())
        logger.warning(f"Error: {e}")