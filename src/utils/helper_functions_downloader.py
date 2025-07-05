from config.loader import load_config
from datetime import datetime
from pathlib import Path
import os
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
import re
import json

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

def login_dsm(driver, wait,LOGIN_URL, USER, PASSWD,):
    """
    Log in to the DSM web interface using credentials from environment variables.
    Maximizes the browser window, navigates to the login page, and submits the login form.

    Parameters:
        driver: Selenium WebDriver instance.
        wait: WebDriverWait instance for waiting on elements.
    """
    driver.maximize_window()
    print("🌐 Navigating to DSM login...")
    driver.get(LOGIN_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    try:
        username_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
        password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
        username_input.clear()
        username_input.send_keys(USER)
        password_input.clear()
        password_input.send_keys(PASSWD + Keys.ENTER)
        print("🔐 DSM login submitted.")
        time.sleep(5)
    except Exception as e:
        print(f"⚠️ Login skipped or already logged in: {e}")



METADATA_FILE = "downloaded_metadata.json"

def go_to_file_station_and_download(driver, wait, target_files):
    """
    Navigate to File Station, download each file in target_files if modified time has changed.
    Also enters the HOURLY folder but only downloads the latest hourly file once.
    """

    ROOT_URL = "https://ithelpdesk-ugml.sg4.quickconnect.to/index.cgi?launchApp=SYNO.SDS.App.FileStation3.Instance"
    HOURLY_URL = (
        "https://ithelpdesk-ugml.sg4.quickconnect.to/index.cgi?"
        "launchApp=SYNO.SDS.App.FileStation3.Instance&"
        "launchParam=openfile%3D%252FV-Optimaise%2520Data%252FBF2%2520AUTO%2520REPORTS%252F"
        "CHARGE%2520AND%2520DUMP%252FHOURLY%252F"
    )

    def normalize(s):
        return re.sub(r'\s+', ' ', s).strip().lower()

    # Load or create metadata file
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r") as f:
            previous_metadata = json.load(f)
    else:
        previous_metadata = {}

    # STEP 1: Root directory
    print("📁 Navigating to File Station root…")
    driver.get(ROOT_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(3)

    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "x-grid3-row")))
        time.sleep(2)
    except:
        print("⚠️ File list did not appear — exiting.")
        return

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
    print(f"📋 {len(file_rows)} items found in root directory.")

    for fname in target_files:
        print(f"🔍 Looking for '{fname}'...")
        matched_row = None
        for row in file_rows:
            if normalize(fname) in normalize(row["name"]):
                matched_row = row
                break

        if not matched_row:
            print(f"⚠️ '{fname}' not found in visible list.")
            continue

        try:
            row_element = matched_row["element"]
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row_element)
            time.sleep(1)

            current_modified = matched_row["modified"]
            previous_modified = previous_metadata.get(fname)

            print(f"📄 Found: {matched_row['name']} | 🕒 Modified: {current_modified} | 📏 Size: {matched_row['size']}")

            if previous_modified and current_modified != previous_modified:
                print(f"📥 Change detected (was: {previous_modified}) → Downloading...")
                ActionChains(driver).move_to_element(row_element).double_click(row_element).perform()
                time.sleep(5)
                print(f"✅ Download complete for '{fname}'")
                previous_metadata[fname] = current_modified
            else:
                print(f"⏩ No change or first-time file — skipping download.")


        except Exception as e:
            print(f"⚠️ Couldn’t download '{fname}': {e}")

    # Save updated metadata
    with open(METADATA_FILE, "w") as f:
        json.dump(previous_metadata, f, indent=2)






    # 2. Go to HOURLY folder
    print("📁 Navigating to HOURLY folder…")
    driver.get(HOURLY_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(5)

    # 3. Sort by Modified Time
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.x-grid3-row")))
        hdr = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//div[contains(@class,'webfm-column-header-text') and text()='Modified Time']")))
        hdr.click()
        time.sleep(1)
        print("✅ Sorted descending.")
    except Exception:
        print("⚠️ Could not sort by Modified Time")

    # 4. Download latest .xlsx file
    try:
        first_row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.x-grid3-row")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_row)

        file_name_el = first_row.find_element(By.CSS_SELECTOR, "div.x-grid3-cell-inner")
        downloaded_file_name = file_name_el.text.strip()
        print(f"📄 File to be downloaded: {downloaded_file_name}")

        ActionChains(driver).move_to_element(first_row).double_click(first_row).perform()
        time.sleep(5)
        print("✅ Latest hourly file download triggered.")

        return downloaded_file_name

    except Exception as e:
        print(f"⚠️ Failed to download latest hourly file: {e}")
        return None










def read_rm_sheet(
    file_path: str,
    RM_SHEET_CONFIG: dict,
    start_date: str = "01-Jun-2025",
    output_dir: str = "outputs"
):
    """
    Reads and combines raw sheets per RM_SHEET_CONFIG (with SINTER averaging).
    Saves a stacked “combined_bunker_data.xlsx” in output_dir.

    Parameters:
        file_path (str): Path to the Excel file.
        RM_SHEET_CONFIG (dict): Configuration for reading sheets.
        start_date (str): Only include rows on/after this date (format: "dd-MMM-yyyy").
        output_dir (str): Directory to save the combined Excel file.
    """
    print(f"\n📄 Reading Excel file: {file_path}")
    start_dt = datetime.strptime(start_date, "%d-%b-%Y").date()
    print(f"   ↪️ Including rows on/after {start_dt.isoformat()}")

    os.makedirs(output_dir, exist_ok=True)
    combined_path = os.path.join(output_dir, "combined_bunker_data.xlsx")

    xls = pd.ExcelFile(file_path)
    parts = []

    for key, cfg in RM_SHEET_CONFIG.items():
        sheet = cfg["sheet_name"]
        cols = cfg["columns"]
        hdr = cfg["header_row"] - 1
        prefix = cfg.get("col_prefix", "")

        if sheet not in xls.sheet_names:
            print(f"⚠️  Sheet '{sheet}' missing, skipping.")
            continue

        print(f"\n🔍  {key}: reading '{sheet}' cols={cols} hdr={hdr+1}")
        df = (pd.read_excel(xls, sheet_name=sheet, usecols=cols, header=hdr)
              .dropna(how="all")
              .reset_index(drop=True))
        if "TIME" in df.columns:
            df = df.drop(columns=["TIME"])

        # SINTER averaging
        if key.upper() == "SINTER" and {"DATE", "SHIFT"}.issubset(df.columns):
            print("   ↪️  SINTER averaging")
            df.columns = df.columns.str.strip()
            if "% T. ALKALI" in df and "Unnamed: 12" in df:
                df = df.rename({"% T. ALKALI": "%Na2O", "Unnamed: 12": "%K2O"}, axis=1)
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.date
            df["SHIFT"] = df["SHIFT"].astype(str).str.strip()

            avg_rows = []
            pairs = [("C-1", "C-2"), ("A-1", "A-2"), ("B-1", "B-2")]
            exclude = ["SHIFT", "BUNKER NO."]
            for dt in df["DATE"].dropna().unique():
                sub = df[df["DATE"] == dt]
                for s1, s2 in pairs:
                    block = sub[sub["SHIFT"].isin([s1, s2])]
                    num = (block.drop(columns=exclude, errors="ignore")
                                 .apply(pd.to_numeric, errors="coerce"))
                    if num.empty: continue
                    r1 = num.iloc[0] if len(num) > 0 else None
                    r2 = num.iloc[1] if len(num) > 1 else None
                    if r1 is not None and r2 is not None:
                        merged = [(v1 + v2) / 2 if pd.notna(v1) and pd.notna(v2) and v1 != 0 and v2 != 0
                                  else (v2 if pd.isna(v1) or v1 == 0 else v1)
                                  for v1, v2 in zip(r1, r2)]
                        out = pd.Series(merged, index=num.columns)
                    elif r1 is not None:
                        out = r1
                    else:
                        out = r2
                    out["DATE"] = dt
                    out["SHIFT"] = s1[0]
                    avg_rows.append(out)
            df = pd.DataFrame(avg_rows)
            if not df.empty:
                df = df[["DATE", "SHIFT"] + [c for c in df.columns if c not in ("DATE", "SHIFT")]]
            else:
                print("   ⚠️  No SINTER averages")

        # Filter by date and ensure DATE and SHIFT are not missing
        if "DATE" in df.columns and "SHIFT" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.date
            df["SHIFT"] = df["SHIFT"].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})
            
            before = len(df)
            df = df[df["DATE"].notna() & df["SHIFT"].notna()]
            df = df[df["DATE"] >= start_dt].reset_index(drop=True)
            
            print(f"   ↪️  Kept {len(df)}/{before} rows with valid DATE and SHIFT ≥ {start_dt}")
        else:
            print("   ⚠️  DATE and SHIFT columns missing — skipping this sheet")
            continue


        # Prefix columns and collect
        df.columns = [prefix + str(c) for c in df.columns]

        if not df.empty:
            parts.append(df)
            print(f"   ✅  {key}: included {len(df)} rows")
        else:
            print(f"   ❌  {key}: no valid data")

    if not parts:
        print("⚠️  No valid data combined — exiting.")
        return

    combined = pd.concat(parts, axis=1)

    # Collapse *_DATE columns to single 'Date'
    date_cols = [c for c in combined.columns if c.upper().endswith("_DATE")]
    if date_cols:
        combined["Date"] = combined[date_cols[0]]
        combined = combined.drop(columns=date_cols)
        combined = combined[["Date"] + [c for c in combined.columns if c != "Date"]]

    # Save to output directory
    combined.to_excel(combined_path, index=False)
    print(f"\n✅  Final combined data written → {combined_path}")







def read_dpr_sheet(file_path, config, output_dir="outputs"):
    """
    Reads DPR sheets based on YAML config, renames fields (no aggregation),
    and saves each processed sheet as a separate Excel file like combined_dpr_<sheet_key>.xlsx
    in the specified output directory.

    Parameters:
        file_path (str): Path to the DPR Excel file.
        config (dict): Configuration dictionary from YAML.
        output_dir (str): Directory to save the combined Excel output.
    """
    dpr_sheets = config["DPR_CONFIG"]["sheets"]
    os.makedirs(output_dir, exist_ok=True)

    wb = load_workbook(file_path, data_only=True)

    for sheet_key, cfg in dpr_sheets.items():
        sheet_name = cfg["sheet_name"]
        date_row = cfg["date_row"] - 1
        col_start, col_end = cfg["date_cols"]
        col_range = range(column_index_from_string(col_start) - 1,
                          column_index_from_string(col_end))

        if sheet_name not in wb.sheetnames:
            print(f"⚠️ Sheet '{sheet_name}' not found — skipping.")
            continue

        ws = wb[sheet_name]

        # Read dates
        dates = [
            cell.value.date() if isinstance(cell.value, datetime) else cell.value
            for cell in [ws.cell(row=date_row + 1, column=col + 1) for col in col_range]
        ]

        # Read raw data rows
        raw_data = {}
        for label, row in cfg["rows"].items():
            raw_data[label] = [
                ws.cell(row=row, column=col + 1).value for col in col_range
            ]

        df = pd.DataFrame({"Date": dates})

        # Rename fields using rename_map
        rename_map = cfg.get("rename_map", {})
        renamed_labels = {v[0]: k for k, v in rename_map.items() if len(v) == 1}

        for original, values in raw_data.items():
            column_name = renamed_labels.get(original, original)
            df[column_name] = values

        # Save a single combined file per sheet
        out_path = os.path.join(output_dir, f"combined_dpr_{sheet_key}.xlsx")
        df.to_excel(out_path, index=False)
        print(f"✅ Saved: {out_path}")





def merge_dpr_and_bunker(
    dpr_path: str,
    bunker_path: str,
    yaml_path: str,
    master_path: str = "master_combined_data.xlsx"
):
    """
    Merges DPR and Bunker Excel files on 'Date', maintains a single master workbook,
    appends only new Date+Shift rows, and enforces presence of Date and Shift.

    Parameters:
        dpr_path (str): Path to the DPR Excel file.
        bunker_path (str): Path to the Bunker Excel file.
        yaml_path (str): Path to YAML config (for fixed_order list).
        master_path (str): Path to the persistent master Excel.
    """


    # 2) Read new DPR & bunker data
    dpr_df    = pd.read_excel(dpr_path)
    bunker_df = pd.read_excel(bunker_path)
    for df in (dpr_df, bunker_df):
        df["Date"]  = pd.to_datetime(df["Date"], errors="coerce").dt.date

    # 3) Combine
    new_df = pd.merge(dpr_df, bunker_df, on="Date", how="outer")

    # 4) Enforce only rows that have both Date and any SHIFT column present
    #    We look for any column ending in '_SHIFT'
    shift_cols = [c for c in new_df.columns if c.endswith("_SHIFT")]
    new_df = new_df.dropna(subset=["Date"] + shift_cols, how="all")

    # 5) Reorder columns per fixed_order, push extras to end
    cols_in_both = [c for c in fixed_order if c in new_df.columns]
    others       = [c for c in new_df.columns if c not in cols_in_both]
    new_df       = new_df[cols_in_both + others]

    # 6) Load existing master (if any) and filter out duplicates
    if os.path.exists(master_path):
        master_df = pd.read_excel(master_path)
        # identify keys: Date + all SHIFT columns
        key_cols   = ["Date"] + shift_cols
        # merge-only truly new rows
        merged = pd.merge(
            new_df, master_df[key_cols].drop_duplicates(),
            on=key_cols, how="left", indicator=True
        )
        new_only = merged[merged["_merge"]=="left_only"].drop(columns="_merge")
        if new_only.empty:
            print("⚠️ No new Date+Shift rows to append.")
            return
        updated = pd.concat([master_df, new_only], ignore_index=True)
    else:
        updated = new_df

    # 7) Save back to master_path
    updated.to_excel(master_path, index=False)
    print(f"✅ Master file updated: {master_path} (added {len(updated) - len(master_df) if 'master_df' in locals() else len(updated)} new rows)")




def merge_hourly_excel(filepath):
    """
    Merge DUMP_REPORT and SH_REPORT by exact DATETIME string match,
    and write the merged result as a new/replaced sheet in the same file.
    """
    try:
        print(f"📂 Reading Excel file: {filepath}")
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names

        if len(sheet_names) < 2:
            print("⚠️ Less than 2 sheets found.")
            return

        # Parse from row 7 (skiprows=6)
        df_dump = xl.parse(sheet_names[0], skiprows=6)
        df_sh = xl.parse(sheet_names[1], skiprows=6)

        # Clean up headers
        df_dump.columns = df_dump.columns.str.strip()
        df_sh.columns = df_sh.columns.str.strip()

        # Drop Unnamed columns (Excel index)
        df_dump = df_dump.loc[:, ~df_dump.columns.str.contains("Unnamed", case=False)]
        df_sh = df_sh.loc[:, ~df_sh.columns.str.contains("Unnamed", case=False)]

        # Treat DATETIME as string
        df_dump['DATETIME'] = df_dump['DATETIME'].astype(str).str.strip()
        df_sh['DATETIME'] = df_sh['DATETIME'].astype(str).str.strip()

        # Merge on DATETIME
        merged_df = pd.merge(df_dump, df_sh, on='DATETIME', how='outer').sort_values('DATETIME')

        # Write back into the same file as sheet "MERGED_EXACT"
        with pd.ExcelWriter(filepath, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            merged_df.to_excel(writer, sheet_name="MERGED_EXACT", index=False)

        print(f"✅ Merged data written to sheet: MERGED_EXACT in {filepath}")
        return filepath

    except Exception as e:
        print(f"❌ Error during merge: {e}")
        return None


    