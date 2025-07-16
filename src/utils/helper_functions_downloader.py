from config.loader import load_config
from datetime import datetime
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

    # STEP 1: ROOT folder download
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

                print("⏳ Waiting 10 seconds for download to complete...")
                time.sleep(0)

                previous_metadata[matched_row["name"]] = current_dt.strftime("%Y-%m-%d %H:%M:%S")
                print(f"✅ Done: {matched_row['name']}")
            else:
                print(f"⏩ No update for '{matched_row['name']}'")
                skipped_files.add(fname)
        except Exception as e:
            print(f"❌ Error downloading {fname}: {e}")
            skipped_files.add(fname)


    # Step 2: HOURLY file based on run_date


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

        # Build target filename
        is_today_mode = "--today" in sys.argv
        today_dt = datetime.today() if is_today_mode else datetime.strptime(run_date, "%d-%b-%Y")
        target_filename = f"CHARGE_AND_DUMP_REPORT_{today_dt.day}_{today_dt.month}_{today_dt.year}.xlsx"
        print(f"🔍 Looking for file: {target_filename}")

        # Pull metadata if in --today mode
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
                        try:
                            file_modified_dt = parse_datetime(modified_str)
                        except:
                            file_modified_dt = datetime.now()

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
    from datetime import datetime
    from openpyxl import load_workbook
    from ruamel.yaml import YAML
    import re
    from calendar import month_abbr
    from config.loader import load_config

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




def read_rm_sheet(
    file_path: str,
    RM_SHEET_CONFIG: dict,
    start_date: str = "11-Jul-2025",
    output_dir: str = "outputs"
):
    """
    Reads and combines raw sheets per RM_SHEET_CONFIG (with SINTER averaging).
    Saves a stacked “combined_bunker_data.xlsx” in output_dir.

    Parameters:
        file_path (str): Path to the Excel file.
        RM_SHEET_CONFIG (dict): Configuration for reading sheets.
        start_date (str or list): Single date (dd-MMM-yyyy) or list of dates.
        output_dir (str): Directory to save the combined Excel file.
    """

    print(f"\n📄 Reading Excel file: {file_path}")

    if isinstance(start_date, list):
        date_list = [datetime.strptime(d, "%d-%b-%Y").date() for d in start_date]
        print(f"   ↪️ Including rows for {len(date_list)} dates: {date_list[0]} to {date_list[-1]}")
    else:
        date_list = [datetime.strptime(start_date, "%d-%b-%Y").date()]
        print(f"   ↪️ Including rows with DATE == {date_list[0]}")

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

        # Filter by date
        if "DATE" in df.columns and "SHIFT" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.date
            df["SHIFT"] = df["SHIFT"].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})

            before = len(df)
            df = df[df["DATE"].notna() & df["SHIFT"].notna()]
            df = df[df["DATE"].isin(date_list)].reset_index(drop=True)

            print(f"   ↪️  Kept {len(df)}/{before} rows matching target dates")
        else:
            print("   ⚠️  DATE and SHIFT columns missing — skipping this sheet")
            continue

        # Prefix columns
        df.columns = [prefix + str(c) for c in df.columns]

        if not df.empty:
            parts.append(df)
            print(f"   ✅  {key}: included {len(df)} rows")
        else:
            print(f"   ❌  {key}: no valid data")
            print(f"📂 Reading RM Excel file: {file_path}")

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

    combined.to_excel(combined_path, index=False)
    print(f"\n✅  Final combined data written → {combined_path}")



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
    out_path = os.path.join(output_dir, "combined_dpr_data.xlsx")
    final_df.to_excel(out_path, index=False)
    print(f"\n✅ Final DPR data written → {out_path}")




def merge_hourly_excel(filepath: str):
    """
    Merge DUMP_REPORT and SH_REPORT by exact DATETIME match.
    Write the merged result to a new Excel file in the working directory.
    """
    try:
        print(f"\n📂 Reading Excel file: {filepath}")
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names

        if len(sheet_names) < 2:
            print("⚠️ Less than 2 sheets found. Expected DUMP_REPORT and SH_REPORT.")
            return None

        df_dump = xl.parse(sheet_names[0], skiprows=6)
        df_sh = xl.parse(sheet_names[1], skiprows=6)

        df_dump.columns = df_dump.columns.str.strip()
        df_sh.columns = df_sh.columns.str.strip()

        df_dump = df_dump.loc[:, ~df_dump.columns.str.contains("Unnamed", case=False)].copy()
        df_sh = df_sh.loc[:, ~df_sh.columns.str.contains("Unnamed", case=False)].copy()

        if "DATETIME" not in df_dump.columns or "DATETIME" not in df_sh.columns:
            print("❌ 'DATETIME' column missing in one of the sheets.")
            return None

        df_dump['DATETIME'] = df_dump['DATETIME'].astype(str).str.strip()
        df_sh['DATETIME'] = df_sh['DATETIME'].astype(str).str.strip()

        print(f"🔄 Merging {len(df_dump)} dump rows and {len(df_sh)} shift rows")

        merged_df = pd.merge(df_dump, df_sh, on='DATETIME', how='outer')
        merged_df = merged_df.sort_values('DATETIME').reset_index(drop=True)

        print(f"✅ Total merged rows: {len(merged_df)}")

        # 🛠️ Save to new file in project directory
        output_path = os.path.join("C:\\Users\\sasik\\Desktop\\evonith_datafeed", "merged_hourly_data.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            merged_df.to_excel(writer, sheet_name="MERGED_EXACT", index=False)

        print(f"✅ Merged data written to: {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error during merge: {e}")
        return None



