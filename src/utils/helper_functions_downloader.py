from config.loader import load_config
from datetime import datetime,timedelta
from pathlib import Path
import os
import sys
import logging
import time
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException
import re
import json
from ruamel.yaml import YAML
from calendar import month_abbr
import unicodedata
from selenium.webdriver.common.action_chains import ActionChains


log = logging.getLogger("root")
project_root = Path(__file__).resolve().parents[2]

# Load config
config = load_config()
fixed_order = config.get("FIXED_COLUMN_ORDER", [])





def extract_datetime_from_filename(filename: str) -> datetime:
    """
    Extracts a datetime object from the filename, assuming the format is %Y_%m_%d_%H_%M_%S.csv.
    
    Parameters:
        filename (str): The filename to extract the datetime from.

    Returns:
        datetime: The parsed datetime object from the filename.
    """
    stem = filename.split('.')[0]
    
    # Parse the datetime string assuming format %Y_%m_%d_%H_%M_%S
    try:
        file_datetime = datetime.strptime(stem, "%Y_%m_%d_%H_%M_%S")
        return file_datetime
    except ValueError as e:
        raise ValueError(f"Filename does not match the expected datetime format: {stem}") from e

def setup_browser_driver():
    """
    Set up and return a Selenium Chrome WebDriver with predefined options.

    Returns:
        webdriver.Chrome: Configured Chrome WebDriver instance.
    """
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def login_eml(driver, wait, LOGIN_URL, USER, PASSWD):
    """
    Log in to the EML web interface using credentials from environment variables.
    Maximizes the browser window, navigates to the login page, and submits the login form.

    Parameters:
        driver: Selenium WebDriver instance.
        wait: WebDriverWait instance for waiting on elements.
    """
    driver.maximize_window()
    print("🌐 Navigating to EML login...")
    driver.get(LOGIN_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    try:
        username_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
        password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
        username_input.clear()
        username_input.send_keys(USER)
        password_input.clear()
        password_input.send_keys(PASSWD + Keys.ENTER)
        print("🔐 EML login submitted.")
        time.sleep(5)
    except Exception as e:
        print(f"⚠️ Login skipped or already logged in: {e}")

METADATA_FILE = "downloaded_metadata.json"

def load_metadata():
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r") as f:
            return json.load(f)
    return {}

def save_metadata(metadata):
    with open(METADATA_FILE, "w") as f:
        json.dump(metadata, f, indent=2)

def parse_datetime(date_str):
    if not date_str or not isinstance(date_str, str):
        return None
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except Exception:
            continue
    return None



def go_to_file_station_and_download(driver, wait, target_files, ROOT_URL, HOURLY_URL, selected_modes, run_date, target_filename=None):
    def normalize(s):
        return re.sub(r'\s+', ' ', s).strip().lower()

    previous_metadata = load_metadata()
    skipped_files = set()
    download_dir = os.path.expanduser("~/Downloads")

    # STEP 1: ROOT folder download (RM + DPR)
    driver.get(ROOT_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(3)

    try:
        wait.until(lambda d: d.find_elements(By.CLASS_NAME, "x-grid3-row"))
        time.sleep(2)
    except:
        print("⚠️ File list did not appear.")
        return skipped_files

    file_rows = driver.execute_script("""
        return Array.from(
            document.querySelectorAll('.x-grid3-body .x-grid3-row')
        ).map(row => {
            const cells = Array.from(row.querySelectorAll('.x-grid3-cell-inner')).map(c => c.innerText.trim());
            return {
                element: row,
                name: cells[0] || "",
                size: cells[1] || "",
                type: cells[2] || "",
                modified: cells[3] || ""
            };
        });
    """)

    mode_file_map = {
        "rm": "11A BF-02 BUNKER",
        "dpr": "BF-02 DPR"
    }

    is_today_mode = "--today" in sys.argv

    for mode in ["rm", "dpr"]:
        if mode not in selected_modes:
            continue

        fname = mode_file_map[mode]
        matched_row = next((row for row in file_rows if normalize(fname) in normalize(row["name"])), None)

        if not matched_row:
            print(f"⚠️ '{fname}' not found.")
            skipped_files.add(fname)
            continue

        try:
            row_element = matched_row["element"]
            current_modified = matched_row["modified"]
            current_dt = parse_datetime(current_modified)

            previous_dt = parse_datetime(previous_metadata.get(matched_row["name"]))

            if not previous_dt or current_dt != previous_dt:
                print(f"📥 Downloading {matched_row['name']}...")
                ActionChains(driver).move_to_element(row_element).double_click(row_element).perform()
                time.sleep(30)
                print(f"✅ Done: {matched_row['name']}")

                # Update metadata
                if mode == "dpr":
                    # Remove other old DPR keys
                    dpr_prefix = "BF-02 DPR"
                    dpr_keys = [k for k in previous_metadata if k.startswith(dpr_prefix)]
                    for k in dpr_keys:
                        if k != matched_row["name"]:
                            del previous_metadata[k]
                previous_metadata[matched_row["name"]] = current_dt.strftime("%Y-%m-%d %H:%M:%S")
                save_metadata(previous_metadata)

            else:
                print(f"⏩ No update for '{matched_row['name']}'")
                skipped_files.add(fname)

        except Exception as e:
            print(f"❌ Error downloading {fname}: {e}")
            skipped_files.add(fname)

    # STEP 2: HOURLY (Charge and Dump)
    if "charge" in selected_modes:
        print("📁 Navigating directly to HOURLY folder…")
        driver.get(HOURLY_URL)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(5)

        try:
            scroll_panel = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "x-grid3-scroller")))
        except Exception:
            print("❌ Could not locate scroll panel.")
            skipped_files.add("charge_and_dump")
            return skipped_files

        # Build filename
        today_dt = datetime.today() if is_today_mode else datetime.strptime(run_date, "%d-%b-%Y")
        target_filename = f"CHARGE_AND_DUMP_REPORT_{today_dt.day}_{today_dt.month}_{today_dt.year}.xlsx"
        print(f"🔍 Looking for file: {target_filename}")

        # Load previous HOURLY_REPORT metadata
        hourly_meta = previous_metadata.get("HOURLY_REPORT", {}) if is_today_mode else {}
        prev_name = hourly_meta.get("name")
        prev_modified = hourly_meta.get("modified")
        prev_dt = parse_datetime(prev_modified) if prev_modified else None

        seen_files = set()
        found = False
        scroll_attempts = 0
        max_scroll_attempts = 50

        while not found and scroll_attempts < max_scroll_attempts:
            rows = driver.find_elements(By.CLASS_NAME, "x-grid3-row")

            for row in rows:
                try:
                    cell = row.find_element(By.CLASS_NAME, "x-grid3-cell-inner")
                    file_name = cell.text.strip()
                    if file_name in seen_files:
                        continue
                    seen_files.add(file_name)

                    if file_name == target_filename:
                        cells = row.find_elements(By.CLASS_NAME, "x-grid3-cell-inner")
                        modified_str = cells[3].text.strip()
                        file_modified_dt = parse_datetime(modified_str) or datetime.now()

                        if is_today_mode and prev_name == target_filename and prev_dt:
                            if file_modified_dt <= prev_dt:
                                print(f"⏩ No update in HOURLY file since last download at {prev_modified}")
                                skipped_files.add("charge_and_dump")
                                return skipped_files

                        print(f"📄 Found file: {file_name} — preparing to download…")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                        time.sleep(1)

                        ActionChains(driver).move_to_element(row).pause(0.5).double_click(row).perform()

                        try:
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "x-window-dlg"))
                            )
                            print("✅ Download confirmation popup detected.")
                        except:
                            print("⚠️ No popup appeared — assuming file is downloading.")

                        print("⏳ Waiting 10 seconds for download to complete...")
                        time.sleep(10)
                        found = True

                        if is_today_mode:
                            previous_metadata["HOURLY_REPORT"] = {
                                "name": target_filename,
                                "modified": file_modified_dt.strftime("%Y-%m-%d %H:%M:%S")
                            }
                            save_metadata(previous_metadata)

                        break
                except Exception as e:
                    print(f"⚠️ Row error: {e}")
                    continue

            if not found:
                ActionChains(driver).move_to_element(scroll_panel).click().send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(1.5)
                scroll_attempts += 1

        if not found:
            print(f"❌ File {target_filename} not found after {scroll_attempts} scroll attempts.")
            skipped_files.add("charge_and_dump")

    return skipped_files


def parse_date_input(start_date):
    if isinstance(start_date, list):
        return [datetime.strptime(d, "%d-%b-%Y").date() for d in start_date]
    return [datetime.strptime(start_date, "%d-%b-%Y").date()]

def normalize_columns(df):
    df.columns = df.columns.str.strip().str.replace(" ", "_").str.upper()
    return df



def average_numeric_group(df, group_cols, skip_cols=[], preserve_order_col="MERGE_KEY"):
    """
    Groups by group_cols, averages numeric columns (ignoring 0s and NaNs),
    discards non-numeric values, and preserves row order.
    """
    results = []

    # Preserve original order
    order_map = (
        df.drop_duplicates(preserve_order_col)
        .reset_index()
        .set_index(preserve_order_col)["index"]
        .to_dict()
    )

    for keys, group in df.groupby(group_cols):
        group = group.drop(columns=skip_cols, errors="ignore")

        avg_row = {}

        # Identify numeric columns reliably
        numeric_cols = []
        for col in group.columns:
            if col in group_cols:
                continue

            # Attempt to convert entire column to numeric (coerce errors to NaN)
            converted = pd.to_numeric(group[col], errors="coerce")
            valid_ratio = converted.notna().sum() / len(group)

            if valid_ratio >= 0.5:
                # Keep column as numeric
                numeric_cols.append(col)

                # Treat 0s as missing
                cleaned = converted.mask(converted == 0, pd.NA)

                # Average ignoring 0 and NaNs
                avg_val = cleaned.mean(skipna=True)
                if pd.notna(avg_val):
                    avg_row[col] = avg_val

        # Add back group keys
        if isinstance(keys, tuple):
            for i, col in enumerate(group_cols):
                avg_row[col] = keys[i]
            merge_key = f"{keys[0]}_{keys[1]}"
        else:
            avg_row[group_cols[0]] = keys
            merge_key = str(keys)

        avg_row["_ORDER"] = order_map.get(merge_key, -1)
        avg_row[preserve_order_col] = merge_key
        results.append(avg_row)

    if not results:
        return pd.DataFrame()

    df_out = pd.DataFrame(results)
    df_out.sort_values("_ORDER", inplace=True)
    df_out.drop(columns=["_ORDER"], inplace=True)
    return df_out

def average_shift_blocks(df):
    """
    Averages rows per (DATE, SHIFT_GROUP), skipping zero/NaNs for numeric cols,
    merges strings, and preserves original SHIFT order.
    """
    df["SHIFT_GROUP"] = df["SHIFT"].str[0].str.upper()
    df["MERGE_KEY"] = df["DATE"].astype(str) + "_" + df["SHIFT_GROUP"]

    # Preserve first-seen MERGE_KEY order
    shift_order = df.drop_duplicates("MERGE_KEY")["MERGE_KEY"].tolist()

    skip_cols = [
        col for col in df.columns if col.upper() in {"SHIFT", "SHIFT_GROUP"}
    ]

    avg_df = average_numeric_group(
        df,
        group_cols=["DATE", "SHIFT_GROUP"],
        skip_cols=skip_cols,
        preserve_order_col="MERGE_KEY"
    )

    if avg_df.empty:
        return pd.DataFrame(columns=["DATE", "SHIFT"])

    # Rename SHIFT_GROUP to SHIFT
    avg_df.rename(columns={"SHIFT_GROUP": "SHIFT"}, inplace=True)

    # Restore original shift order using preserved list
    avg_df["MERGE_KEY"] = avg_df["DATE"].astype(str) + "_" + avg_df["SHIFT"]
    avg_df["_ORDER"] = avg_df["MERGE_KEY"].apply(lambda x: shift_order.index(x) if x in shift_order else -1)
    avg_df = avg_df.sort_values("_ORDER").drop(columns=["MERGE_KEY", "_ORDER"]).reset_index(drop=True)

    # Reorder columns to put DATE and SHIFT first
    cols = ["DATE", "SHIFT"] + [col for col in avg_df.columns if col not in {"DATE", "SHIFT"}]
    avg_df = avg_df[cols]

    print("✅ AVG is Done")
    return avg_df


def read_excel_sheet(xls, sheet, cols, hdr):
    if sheet not in xls.sheet_names:
        print(f"⚠  Sheet '{sheet}' missing, skipping.")
        return None

    df = pd.read_excel(xls, sheet_name=sheet, usecols=cols, header=hdr).dropna(how="all").reset_index(drop=True)

    # Normalize column names (optional)
    df.columns = [str(col).strip().upper() for col in df.columns]

    # Define always-keep-as-text columns
    always_text_cols = ['DATE', 'SHIFT', 'SOURCE', 'BUNKER','BNK NO', 'ONLINE/OFFLINE','SORUCE']

    for col in df.columns:
        if any(key in col for key in always_text_cols):
            continue
        # Try to coerce to float; replace non-numeric with NaN
        df[col] = pd.to_numeric(df[col], errors='coerce')

    return df.drop(columns=["TIME"], errors="ignore")

def filter_by_date_and_shift(df, date_list, logger=None):
    col_map = {col.strip().upper(): col for col in df.columns}
    date_col = col_map.get("DATE")
    shift_col = next((col_map[k] for k in col_map if "SHIFT" in k), None)
    if not date_col or not shift_col:
        msg = "⚠  Missing DATE or SHIFT column — skipping this sheet"
        print(msg,df) if logger is None else logger.warning(msg)
        return None
    parsed = pd.to_datetime(df[date_col], format="%d-%m-%Y", errors="coerce")
    if parsed.isna().all():
        parsed = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    df[date_col] = parsed.dt.date
    df[shift_col] = df[shift_col].astype(str).str.strip().str.upper()
    VALID_SHIFTS = {"A", "B", "C"}
    df[shift_col] = df[shift_col].apply(
        lambda x: x[0] if x and isinstance(x, str) and x[0] in VALID_SHIFTS else pd.NA
    )
    before = len(df)
    df = df[df[date_col].isin(date_list) & df[shift_col].notna()].reset_index(drop=True)
    after = len(df)
    # msg = f"   ↪  Kept {after}/{before} rows with valid DATE and SHIFT in {VALID_SHIFTS}"
    # print(msg) if logger is None else logger.info(msg)
    return df

def split_online_offline_and_merge(df):
    df = normalize_columns(df)
    df["ONLINE/OFFLINE"] = df["ONLINE/OFFLINE"].astype(str).str.strip()
    df["SHIFT"] = df["SHIFT"].astype(str).str.strip()
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce", dayfirst=True).dt.date
    if "BNK_NO" in df.columns:
        df["BNK_NO"] = df["BNK_NO"].astype(str).str.strip()
    df["MERGE_KEY"] = df["DATE"].astype(str) + "_" + df["SHIFT"]
    df["IS_ONLINE"] = df["ONLINE/OFFLINE"].str.contains(r"(?:ONLINE|EML)", case=False, na=False)

    def reduce(g, suffix):
        base = {col: g[col].iloc[0] for col in ["MERGE_KEY", "DATE", "SHIFT"] if col in g.columns}
        numeric_cols = []
        num_data = {}
        for col in g.columns:
            if col in base or col in ["IS_ONLINE", "MERGE_KEY", "DATE", "SHIFT"]:
                continue
            try:
                converted = pd.to_numeric(g[col], errors="coerce")
                if converted.notna().sum() >= 0.8 * len(g):
                    vals = converted.replace(0, pd.NA).dropna()
                    num_data[col + suffix] = round(vals.mean(skipna=True), 3) if not vals.empty else pd.NA
                    numeric_cols.append(col)
            except:
                continue
        for col in g.columns:
            if col in base or col in numeric_cols or col in ["IS_ONLINE"]:
                continue
            unique_vals = g[col].dropna().astype(str).str.strip().unique()
            base[col + suffix] = " | ".join(sorted(set(unique_vals)))
        base.update(num_data)
        return base

    records = []
    for (key, shift), group in df.groupby(["MERGE_KEY", "SHIFT"]):
        online, offline = group[group["IS_ONLINE"]], group[~group["IS_ONLINE"]]
        if len(online) > 1:
            print(f"\n🟦 Multiple ONLINE rows for DATE: {group['DATE'].iloc[0]}, SHIFT: {shift}")
            print(online.to_string(index=False))
        if len(offline) > 1:
            print(f"\n🟥 Multiple OFFLINE rows for DATE: {group['DATE'].iloc[0]}, SHIFT: {shift}")
            print(offline.to_string(index=False))
        combined = {}
        if not online.empty: combined.update(reduce(online, "_ON"))
        if not offline.empty: combined.update(reduce(offline, "_OFF"))
        records.append(combined)

    df_out = pd.DataFrame(records)
    final_df = pd.merge(df.drop_duplicates("MERGE_KEY")[["MERGE_KEY"]].reset_index(), df_out, on="MERGE_KEY")
    return final_df.drop(columns=["MERGE_KEY", "index"])

def prefix_columns(df, prefix):
    df.columns = [prefix + str(c) for c in df.columns]
    return df

def write_combined_file(parts, combined_path, date_list):
    import os
    import pandas as pd

    SHIFT_ORDER = ["C", "A", "B"]
    SHIFT_TIME = {"A": "07:00", "B": "15:00", "C": "1:00"}

    # Combine all dataframes
    combined = pd.concat(parts, axis=1)

    # Find all _DATE columns
    date_cols = [c for c in combined.columns if c.upper().endswith("_DATE")]

    if not date_cols:
        raise ValueError("No _DATE columns found in the combined dataframe.")

    # Take the first _DATE column as the main Date
    combined["Date"] = pd.to_datetime(combined[date_cols[0]], errors="coerce")

    # Drop all _DATE columns
    combined.drop(columns=date_cols, inplace=True, errors="ignore")

    # Assign cyclic shift order: C, A, B, C, A, B, ...
    combined["SHIFT"] = [SHIFT_ORDER[i % 3] for i in range(len(combined))]

    # Replace date with shift-adjusted datetime
    combined["Date"] = combined.apply(
        lambda r: pd.to_datetime(f"{r['Date'].date()} {SHIFT_TIME[r['SHIFT']]}"),
        axis=1
    )

    # Drop SHIFT column
    combined.drop(columns="SHIFT", inplace=True)

    # Reorder columns to have Date first
    combined = combined[["Date"] + [col for col in combined.columns if col != "Date"]]

    # Filter existing Excel to avoid duplicates when date_list > 1
    target_dates = set(pd.to_datetime(date_list).date)
    if len(date_list) > 1 and os.path.exists(combined_path):
        existing = pd.read_excel(combined_path)
        existing["Date"] = pd.to_datetime(existing["Date"], errors="coerce")
        existing = existing[~existing["Date"].dt.date.isin(target_dates)].reset_index(drop=True)
        combined = pd.concat([existing, combined], ignore_index=True)

    # Write final file
    combined.to_excel(combined_path, index=False)

def read_rm_sheet(file_path, RM_SHEET_CONFIG, start_date="11-Jul-2025", output_dir="outputs"):
    print(f"\n📄 Reading Excel file: {file_path}")
    date_list = parse_date_input(start_date)
    print(f"   ↪ Including rows from {date_list[0]} to {date_list[-1]}" if len(date_list) > 1 else f"   ↪ Including rows with DATE == {date_list[0]}")
    os.makedirs(output_dir, exist_ok=True)
    combined_path = os.path.join(output_dir, "combined_bunker_data.xlsx")
    xls = pd.ExcelFile(file_path)
    parts = []

    for key, cfg in RM_SHEET_CONFIG.items():
        sheet, cols, hdr = cfg["sheet_name"], cfg["columns"], cfg["header_row"] - 1
        prefix = cfg.get("col_prefix", "")
        df = read_excel_sheet(xls, sheet, cols, hdr)
        if df is None:
            continue

        df = normalize_columns(df)
        df = filter_by_date_and_shift(df, date_list)
        if df is None or df.empty:
            print(f"   ❌  {key}: no valid data")
            
            continue

        if "ONLINE/OFFLINE" in df.columns:
            df = split_online_offline_and_merge(df)
            # print(f"   ✅  {key}: merged ONLINE + OFFLINE per SHIFT+DATE")

        # ✅ Apply averaging if multiple rows for a normalized shift group
        if "SHIFT" in df.columns and "DATE" in df.columns:
            shift_counts = df.groupby(["DATE", "SHIFT"]).size().reset_index(name="count")
            if any(shift_counts["count"] > 1):
                df = average_shift_blocks(df)
                
                print(f"   🔄  {key}: averaged multiple rows per SHIFT block")

        df = prefix_columns(df, prefix)
        parts.append(df)

    if parts:
        write_combined_file(parts, combined_path, date_list)
    else:
        print("⚠  No valid data combined — exiting.")




# Function to parse sheet date from name like "Jun'25"
def parse_sheet_date(name):
    m = re.match(r"([A-Za-z]+)'(\d{2})$", name)
    if not m:
        return None
    try:
        return (2000 + int(m[2]), list(month_abbr).index(m[1][:3].title()))
    except:
        return None


def update_dpr_config_from_excel(excel_path, yaml_path, run_date):


    config = load_config(yaml_path)
    yaml = YAML()
    yaml.preserve_quotes = True

    sheets = config.get("DPR_CONFIG", {}).get("sheets", {})
    if not sheets:
        raise RuntimeError("Missing DPR_CONFIG.sheets in YAML")

    # convert run_date to datetime.date
    run_dt = datetime.strptime(run_date, "%d-%b-%Y").date()
    target_sheet = None
    target_dt = (run_dt.year, run_dt.month)

    # match sheet by run date month and year
    wb = load_workbook(excel_path, data_only=True)
    dated = [(s, parse_sheet_date(s)) for s in wb.sheetnames if parse_sheet_date(s)]

    for sname, parsed in dated:
        if parsed and (parsed[0], parsed[1]) == target_dt:
            target_sheet = sname
            break

    if not target_sheet:
        print(f"⚠️ No matching sheet found for run date {run_date}. Using latest available sheet.")
        target_sheet, _ = max(dated, key=lambda x: x[1])

    sheet_key = target_sheet.replace("'", "")
    ws = wb[target_sheet]

    if sheet_key not in sheets:
        old_key = next(iter(sheets))
        sheets[sheet_key] = sheets.pop(old_key)
        print(f"ℹ️ Renamed YAML month block '{old_key}' ➔ '{sheet_key}'")

    block = sheets[sheet_key]
    block["sheet_name"] = target_sheet

    old_rows = block.get("rows", {})
    found_rows = {}

    for r in range(1, ws.max_row + 1):
        texts = [str(ws.cell(r, c).value).strip() for c in range(1, 8) if ws.cell(r, c).value]
        for label in old_rows:
            if label in texts and label not in found_rows:
                found_rows[label] = r

    block["rows"] = {k: found_rows[k] for k in old_rows if k in found_rows}
    for k in old_rows:
        if k not in found_rows:
            print(f"⚠️ '{k}' not found in sheet '{target_sheet}'")

    with open(yaml_path, "w") as f:
        yaml.dump(config, f)

    print(f"✅ Updated '{yaml_path}' with sheet '{target_sheet}' → rows: {block['rows']}")


def read_dpr_sheet(
    file_path: str,
    config: dict,
    start_date: str = "11-Jun-2025",
    output_dir: str = "outputs"
):
    """
    Reads DPR sheets based on YAML config, renames fields, filters by start_date,
    and saves combined output to a single fixed file: combined_dpr_data.xlsx
    """
    from openpyxl.utils import column_index_from_string
    import pandas as pd
    from openpyxl import load_workbook
    import os

    dpr_sheets = config["DPR_CONFIG"]["sheets"]
    os.makedirs(output_dir, exist_ok=True)
    start_dt = datetime.strptime(start_date, "%d-%b-%Y").date()
    wb = load_workbook(file_path, data_only=True)

    all_parts = []

    print(f"\n📘 Reading DPR Excel file: {file_path}")
    print(f"📅 Filtering for date: {start_dt}")

    for sheet_key, cfg in dpr_sheets.items():
        sheet_name = cfg["sheet_name"]
        date_row = cfg["date_row"] - 1
        col_start, col_end = cfg["date_cols"]
        col_range = range(
            column_index_from_string(col_start) - 1,
            column_index_from_string(col_end)
        )

        if sheet_name not in wb.sheetnames:
            print(f"⚠️ Sheet '{sheet_name}' not found — skipping.")
            continue

        ws = wb[sheet_name]

        # Step 1: Read date headers
        raw_dates = [ws.cell(row=date_row + 1, column=col + 1).value for col in col_range]
        parsed_dates = []
        valid_indices = []

        for i, val in enumerate(raw_dates):
            parsed = pd.to_datetime(val, errors="coerce")
            if pd.notna(parsed):
                parsed_dates.append(parsed.date())
                valid_indices.append(i)

        if not parsed_dates:
            print(f"⚠️ No valid dates found in sheet '{sheet_name}' — skipping.")
            continue

        # Step 2: Read rows for this sheet
        raw_data = {}
        for label, row in cfg["rows"].items():
            all_values = [ws.cell(row=row, column=col + 1).value for col in col_range]
            filtered_values = [all_values[i] for i in valid_indices]
            raw_data[label] = filtered_values

        # Step 3: Drop columns that are empty across all rows
        non_empty_mask = [
            any(raw_data[label][i] is not None for label in raw_data)
            for i in range(len(valid_indices))
        ]
        filtered_dates = [parsed_dates[i] for i, keep in enumerate(non_empty_mask) if keep]

        if not filtered_dates:
            print(f"⚠️ No data found for sheet '{sheet_name}' on any date — skipping.")
            continue

        # Step 4: Build DataFrame for this sheet
        df = pd.DataFrame({"Date": filtered_dates})
        rename_map = cfg.get("rename_map", {})
        reverse_map = {v[0]: k for k, v in rename_map.items() if isinstance(v, list) and len(v) == 1}

        for original, all_values in raw_data.items():
            values = [all_values[i] for i, keep in enumerate(non_empty_mask) if keep]
            colname = reverse_map.get(original, original)
            df[colname] = values
            # Derive TOTAL_COKE_MT as sum of COKE_ONLINE_MT and COKE_OFFLINE_MT
        if "COKE_ONLINE_MT" in df.columns and "COKE_OFFLINE_MT" in df.columns:
            df["TOTAL_COKE_MT"] = df["COKE_ONLINE_MT"].fillna(0) + df["COKE_OFFLINE_MT"].fillna(0)


        # Step 5: Filter for only start_date
        before = len(df)
        df = df[df["Date"] == start_dt].reset_index(drop=True)
        after = len(df)

        print(f"🔍 Sheet '{sheet_name}': kept {after}/{before} rows for {start_dt}")

        if not df.empty:
            # df.insert(0, "Sheet", sheet_key)
            all_parts.append(df)

    # Final save
    if not all_parts:
        print("⚠️ No DPR data found — nothing to write.")
        return

    final_df = pd.concat(all_parts, ignore_index=True)
    os.path.join(output_dir, "combined_dpr_data.xlsx")
    print(f"\n✅ DPR data extracted for {start_date}")
    return final_df


# def merge_hourly_excel(filepath: str):
#     """
#     Merge DUMP_REPORT and SH_REPORT by exact DATETIME match.
#     Write the merged result to a new Excel file in the working directory.
#     """
#     try:
#         print(f"\n📂 Reading Excel file: {filepath}")
#         xl = pd.ExcelFile(filepath)
#         sheet_names = xl.sheet_names

#         if len(sheet_names) < 2:
#             print("⚠️ Less than 2 sheets found. Expected DUMP_REPORT and SH_REPORT.")
#             return None

#         df_dump = xl.parse(sheet_names[0], skiprows=6)
#         df_sh = xl.parse(sheet_names[1], skiprows=6)

#         df_dump.columns = df_dump.columns.str.strip()
#         df_sh.columns = df_sh.columns.str.strip()

#         df_dump = df_dump.loc[:, ~df_dump.columns.str.contains("Unnamed", case=False)].copy()
#         df_sh = df_sh.loc[:, ~df_sh.columns.str.contains("Unnamed", case=False)].copy()

#         if "DATETIME" not in df_dump.columns or "DATETIME" not in df_sh.columns:
#             print("❌ 'DATETIME' column missing in one of the sheets.")
#             return None

#         df_dump['DATETIME'] = df_dump['DATETIME'].astype(str).str.strip()
#         df_sh['DATETIME'] = df_sh['DATETIME'].astype(str).str.strip()

#         print(f"🔄 Merging {len(df_dump)} dump rows and {len(df_sh)} shift rows")

#         merged_df = pd.merge(df_dump, df_sh, on='DATETIME', how='outer')
#         merged_df = merged_df.sort_values('DATETIME').reset_index(drop=True)

#         print(f"✅ Total merged rows: {len(merged_df)}")

#         # 🛠️ Save to new file in project directory
#         output_path = os.path.join("C:\\Users\\sasik\\Desktop\\evonith_datafeed", "merged_hourly_data.xlsx")

#         with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
#             merged_df.to_excel(writer, sheet_name="MERGED_EXACT", index=False)

#         print(f"✅ Merged data written to: {output_path}")
#         return output_path

#     except Exception as e:
#         print(f"❌ Error during merge: {e}")
#         return None

def process_shiftwise_charge_data(file_today, file_yesterday, output_dir, run_date_str):
    def assign_shift(dt, ref_date):
        if pd.isnull(dt):
            return None
        if dt.date() == ref_date.date() and dt.hour >= 23:
            return None  # Exclude 23:00–23:59 of current day
        if (dt.hour >= 23 and dt.date() < ref_date.date()) or dt.hour < 7:
            return 'C'
        elif 7 <= dt.hour < 15:
            return 'A'
        elif 15 <= dt.hour < 23:
            return 'B'

    dfs = []
    for path in [file_yesterday, file_today]:
        if path and os.path.exists(path):
            xl = pd.ExcelFile(path)
            sh_sheet = next((s for s in xl.sheet_names if "SH" in s.upper()), None)
            if sh_sheet:
                df = xl.parse(sh_sheet, skiprows=6)
                df.columns = df.columns.str.strip()
                df = df.loc[:, ~df.columns.str.contains("Unnamed", case=False)]
                df = df.drop(columns=[c for c in df.columns if "CHARGE_NO" in c.upper()], errors="ignore")
                df["DATETIME"] = pd.to_datetime(df["DATETIME"], errors='coerce')
                dfs.append(df)

    if not dfs:
        print("❌ No charge data files found.")
        return

    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df = combined_df[combined_df["DATETIME"].notna()].copy()

    target_date = datetime.strptime(re.search(r"(\d{1,2})-(\w+)-(\d{4})", run_date_str).group(0), "%d-%b-%Y")
    start_time = target_date - timedelta(hours=1)
    end_time = target_date + timedelta(days=1)

    print(f"🕒 Filtering data from {start_time} to {end_time}")
    combined_df = combined_df[(combined_df["DATETIME"] >= start_time) & (combined_df["DATETIME"] < end_time)].copy()
    combined_df["SHIFT"] = combined_df["DATETIME"].apply(lambda x: assign_shift(x, target_date))

    if combined_df.empty:
        print("⚠️ No data in selected range.")
        return

    print("📊 Hourly data counts:")
    print(combined_df["DATETIME"].dt.hour.value_counts().sort_index())

    print("📊 Shift-wise row counts:")
    print(combined_df["SHIFT"].value_counts())

    print("\n📅 Shift-wise timestamps:")
    for shift in ["C", "A", "B"]:
        timestamps = combined_df.loc[combined_df["SHIFT"] == shift, "DATETIME"].dt.strftime("%Y-%m-%d %H:%M:%S").tolist()
        print(f"  Shift {shift} → {len(timestamps)} entries")
        # if timestamps:
        #     print(timestamps)

    # ➕ Only sum columns containing 'ACT' in their name
    act_columns = [col for col in combined_df.columns if "ACT" in col.upper()]
    shift_sums = combined_df.groupby("SHIFT")[act_columns].sum(numeric_only=True).reset_index()

    output_path = os.path.join(output_dir, f"summed_shift_values_{target_date.strftime('%Y_%m_%d')}.xlsx")
    shift_sums.to_excel(output_path, index=False)
    print(f"\n✅ Shiftwise ACT-only charge report written → {output_path}")
