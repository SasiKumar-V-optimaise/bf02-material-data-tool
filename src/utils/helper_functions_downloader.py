from config.loader import load_config
from datetime import datetime
from pathlib import Path
import os
import logging
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

log = logging.getLogger("root")
project_root = Path(__file__).resolve().parents[2]

# Load config
config = load_config()

# Load credentials and URLs
load_dotenv()
LOGIN_URL = os.getenv("EVONITH_URL")
FILESTATION_URL = os.getenv("EVONITH_FILESTATION_URL")
USER = os.getenv("EVONITH_USER")
PASSWD = os.getenv("EVONITH_PASS")

def clean_folder(folder_path: str):
    """
    Deletes all .csv files in the specified folder.

    Parameters:
        folder_path (str): The path to the folder where .csv files need to be deleted.
    """
    folder = Path(folder_path)
    
    # Check if the folder exists
    if not folder.is_dir():
        raise ValueError(f"The specified path is not a directory: {folder_path}")
    
    # Iterate through all .csv files in the folder and delete them
    csv_files = folder.glob("*.csv")  # Matches all .csv files in the folder
    deleted_files = []
    for csv_file in csv_files:
        csv_file.unlink()  # Deletes the file
        deleted_files.append(csv_file.name)
    
    if deleted_files:
        log.info(f"Deleted the following CSV files: {', '.join(deleted_files)}")
    else:
        log.info("No CSV files found to delete.")

def rename_file(file: str) -> None:
  os.rename(project_root / "data" / file, project_root / "data" / "data.csv")

def extract_datetime_from_filename(filename: str) -> datetime:
    """
    Extracts a datetime object from the filename, assuming the format is %Y_%m_%d_%H_%M_%S.csv.
    
    Parameters:
        file_path (str): The full path to the file.

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
    """
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def login_dsm(driver, wait):
    """
    Log in to the DSM web interface using credentials from environment variables.
    Maximizes the browser window, navigates to the login page, and submits the login form.
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

def go_to_file_station_and_download(driver, wait):
    """
    Navigate to the File Station page and attempt to download a specific file by double-clicking it.
    Handles errors and allows for manual intervention if needed.
    """
    print("📁 Navigating to File Station...")
    driver.get(FILESTATION_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    print("✅ Base page HTML loaded.")
    try:
        print("⏳ Waiting for file grid to load...")
        time.sleep(5)  # Give time for the grid to load
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        keyword = "11A BF-02 BUNKER  2025-26.xlsx"
        # TODO: Make this keyword configurable
        file_xpath = f"//div[contains(@class, 'webfm-file-type-icon') and contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword.lower()}')]"
        print(f"🔍 Searching grid for file containing: {keyword}")
        file_name_el = wait.until(EC.presence_of_element_located((By.XPATH, file_xpath)))
        file_wrap_el = file_name_el.find_element(By.XPATH, "..")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_wrap_el)
        wait.until(EC.visibility_of(file_wrap_el))
        time.sleep(1)
        ActionChains(driver).move_to_element(file_wrap_el).pause(1).double_click(file_wrap_el).perform()
        print("✅ File double-clicked. Download should begin.")
        time.sleep(10)
    except Exception as e:
        print(f"⚠️ Error during file interaction: {e.__class__.__name__}: {str(e)}")
        driver.save_screenshot("error_debug.png")
        try:
            print("🧾 Elements matching 'bunker':")
            matches = driver.find_elements(By.XPATH, f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword.lower()}')]")
            for m in matches:
                print(m.get_attribute("outerHTML"))
        except:
            print("⚠️ Could not print matching elements.")
        input("⏸️ Manual intervention needed. Press Enter to continue...")

def read_sheets_by_config(file_path, sheet_configs, start_date="01-Jun-2025"):
    """
    Read and process multiple sheets from an Excel file according to the provided configuration.
    Filters rows by start date, applies sheet-specific logic (e.g., SINTER averaging), and writes the combined result to an output file.
    
    Args:
        file_path (str): Path to the Excel file.
        sheet_configs (dict): Configuration for each sheet (columns, header row, etc.).
        start_date (str): Only include rows on or after this date (format: 'dd-MMM-yyyy').
    """
    print(f"\n📄 Reading Excel file: {file_path}")
    start_date_obj = datetime.strptime(start_date, "%d-%b-%Y").date()
    print(f"   ↪️  Will include rows on/after {start_date_obj.isoformat()}")
    writer = pd.ExcelWriter("combined_output_sheety.xlsx", engine="openpyxl")
    try:
        xls = pd.ExcelFile(file_path)
        prepared = []
        for sheet_name, cfg in sheet_configs.items():
            cols       = cfg["columns"]
            header_row = cfg["header_row"] - 1
            prefix     = cfg.get("col_prefix", "")
            if sheet_name not in xls.sheet_names:
                print(f"⚠️ Sheet '{sheet_name}' not found — skipping.")
                continue
            print(f"\n🔍 Processing '{sheet_name}' (cols={cols}, header_row={header_row+1})")
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=cols, header=header_row)
            df = df.dropna(how="all").reset_index(drop=True)
            if "TIME" in df.columns:
                df = df.drop(columns=["TIME"])
            if sheet_name.upper() == "SINTER":
                print("   ↪️  SINTER: doing shift‑wise averaging …")
                df.columns = [str(c).strip() for c in df.columns]
                if "% T. ALKALI" in df.columns and "Unnamed: 12" in df.columns:
                    df = df.rename({"% T. ALKALI": "%Na2O", "Unnamed: 12": "%K2O"}, axis=1)
                if {"DATE", "SHIFT"}.issubset(df.columns):
                    df["DATE"]  = pd.to_datetime(df["DATE"], errors="coerce").dt.date
                    df["SHIFT"] = df["SHIFT"].astype(str).str.strip()
                    pairs   = [("C-1","C-2"), ("A-1","A-2"), ("B-1","B-2")]
                    exclude = ["SHIFT","BUNKER NO."]
                    avg_rows = []
                    for dt in df["DATE"].dropna().unique():
                        day = df[df["DATE"] == dt]
                        for s1, s2 in pairs:
                            block = day[day["SHIFT"].isin([s1, s2])]
                            num   = block.drop(columns=exclude, errors="ignore")\
                                         .apply(pd.to_numeric, errors="coerce")
                            if num.empty: 
                                continue
                            row1 = num.iloc[0] if len(num) >= 1 else None
                            row2 = num.iloc[1] if len(num) >= 2 else None
                            if row1 is not None and row2 is not None:
                                vals = []
                                for v1, v2 in zip(row1, row2):
                                    if pd.isna(v1) or v1 == 0:
                                        vals.append(v2)
                                    elif pd.isna(v2) or v2 == 0:
                                        vals.append(v1)
                                    else:
                                        vals.append((v1 + v2) / 2)
                                out = pd.Series(vals, index=num.columns)
                            elif row1 is not None:
                                out = row1
                            elif row2 is not None:
                                out = row2
                            else:
                                continue
                            out["SHIFT"] = s1[0]
                            out["DATE"]  = dt
                            avg_rows.append(out)
                    df = pd.DataFrame(avg_rows)
                    if not df.empty:
                        cols0 = df.columns.tolist()
                        df = df[["DATE","SHIFT"] + [c for c in cols0 if c not in ("DATE","SHIFT")]]
                    else:
                        print("   ⚠️  SINTER: no averaged data generated")
                else:
                    print("   ⚠️  SINTER: missing DATE or SHIFT — skipping averaging")
            date_cols = [c for c in df.columns if c.strip().upper() == "DATE"]
            if date_cols:
                dc = date_cols[0]
                df[dc] = pd.to_datetime(df[dc], errors="coerce").dt.date
                before = len(df)
                df = df[df[dc] >= start_date_obj].reset_index(drop=True)
                print(f"   ↪️  Kept {len(df)}/{before} rows on/after {start_date_obj.isoformat()}")
            else:
                print(f"   ⚠️  '{sheet_name}' has no DATE column — including all rows")
            df.columns = [f"{prefix}{c}" for c in df.columns]
            df = df.reset_index(drop=True)
            if not df.empty:
                prepared.append(df)
                print(f"   ✅  included {len(df)} rows from '{sheet_name}'")
            else:
                print(f"   ❌  no rows on/after {start_date_obj.isoformat()} for '{sheet_name}'")
        if prepared:
            combined = pd.concat(prepared, axis=1)
            combined.to_excel(writer, "Combined", index=False)
            print("\n✅  Written combined sheet to 'combined_output_sheet.xlsx'")
        else:
            print("\n⚠️  nothing to write — all sheets empty or filtered out")
    except Exception as e:
        print(f"\n❌  Fatal error: {e}")
    finally:
        writer.close()


