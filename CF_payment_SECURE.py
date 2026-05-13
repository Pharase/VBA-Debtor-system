"""
CF_payment_SECURE.py - Secure version with all hardcoded credentials removed
================================================================================
⚠️ SECURITY NOTICE:
- All hardcoded email credentials have been REMOVED
- All hardcoded file paths have been replaced with environment variables
- Network shares are configurable

REQUIRED ENVIRONMENT VARIABLES:
- EMAIL_ADDRESS: Your email address
- EMAIL_PASSWORD: Your email password (use app-specific password)
- DOWNLOAD_DIR: Directory for downloaded files
- SUMMARY_REPORT_DIR: Directory for summary reports
- OUTPUT_DIR: Directory for output files

BEFORE DEPLOYING:
1. Create .env file with required variables
2. Use app-specific password for Office 365
3. Use OAuth 2.0 authentication instead (recommended)
================================================================================
"""

import os
import sys
from pathlib import Path
from datetime import datetime
import time
import glob
import re

try:
    from dotenv import load_dotenv
except ImportError:
    print("⚠️ WARNING: python-dotenv not installed. Using system environment variables only.")

# Load environment variables
env_file = Path(__file__).parent / ".env"
if env_file.exists():
    load_dotenv(env_file)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd


def get_credentials():
    """
    Retrieve email credentials from secure storage.
    
    Returns:
        tuple: (email, password)
        
    Raises:
        EnvironmentError: If credentials not set
    """
    email = os.getenv('EMAIL_ADDRESS')
    password = os.getenv('EMAIL_PASSWORD')
    
    if not email or not password:
        raise EnvironmentError(
            "❌ Missing email credentials!\n"
            "Set environment variables:\n"
            "  EMAIL_ADDRESS\n"
            "  EMAIL_PASSWORD (use app-specific password for Office 365)\n"
            "\n⚠️ BETTER: Migrate to OAuth 2.0 / Microsoft Graph API"
        )
    
    return email, password


def get_paths():
    """
    Retrieve file paths from environment variables.
    
    Returns:
        dict: Dictionary with all required paths
    """
    paths = {
        'download_dir': os.getenv('DOWNLOAD_DIR'),
        'summary_report_dir': os.getenv('SUMMARY_REPORT_DIR'),
        'output_dir': os.getenv('OUTPUT_DIR'),
    }
    
    missing = [k for k, v in paths.items() if not v]
    if missing:
        raise EnvironmentError(
            f"❌ Missing environment variables: {', '.join(missing)}\n"
            "Set these in your .env file or environment."
        )
    
    return paths


def wait_and_rename(download_dir, new_filename, timeout=30):
    """
    Wait for file download to complete and rename it.
    """
    print("⏳ Waiting for download to finish...")
    end_time = time.time() + timeout
    downloaded_file = None

    while time.time() < end_time:
        files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload")]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
            downloaded_file = files[0]
            break
        time.sleep(1)

    if downloaded_file:
        old_path = os.path.join(download_dir, downloaded_file)
        new_path = os.path.join(download_dir, new_filename)
        os.rename(old_path, new_path)
        print(f"✅ File renamed to: {new_filename}")
        return new_path
    else:
        print("⚠️ Download did not finish in time.")
        return None


def setup_browser(download_dir):
    """
    Configure Chrome browser with security best practices.
    """
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    return webdriver.Chrome(service=Service(), options=options)


def download_emails():
    """
    Download payment confirmation emails using Selenium.
    
    ⚠️ SECURITY WARNING: This uses password authentication.
    RECOMMENDATION: Use Microsoft Graph API with OAuth 2.0 instead.
    """
    try:
        email, password = get_credentials()
    except EnvironmentError as e:
        print(f"\n{e}")
        return None
    
    try:
        paths = get_paths()
    except EnvironmentError as e:
        print(f"\n{e}")
        return None
    
    download_dir = paths['download_dir']
    os.makedirs(download_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d")
    check_load_output_path = os.path.join(download_dir, f"Confirm_Payment_{timestamp}.xlsx")
    
    if Path(check_load_output_path).exists():
        print(f"✅ File already exists: {check_load_output_path}")
        return check_load_output_path
    
    driver = setup_browser(download_dir)
    
    try:
        driver.get("https://outlook.office.com/mail/inbox")
        wait = WebDriverWait(driver, 20)
        
        # Login
        print("🔐 Logging in...")
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]'))).send_keys(email)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
        time.sleep(3)
        
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
        time.sleep(3)
        
        # Handle stay signed in dialog
        try:
            wait.until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
        except:
            pass
        
        time.sleep(3)
        
        # Search emails
        print("🔍 Searching for emails...")
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchBoxId-Mail"]'))).click()
        
        wait.until(EC.presence_of_element_located((By.ID, 'filtersButtonId'))).click()
        time.sleep(1)
        
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Subject-ID"]')))
        search_box.click()
        search_box.send_keys('Confirm Payment')
        
        # Date filter
        current_date = datetime.now().strftime("%m/%d/%Y")
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Select a date"]')))
        date_input.click()
        date_input.send_keys(current_date)
        
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "ms-Button--primary")]')))
        search_button.click()
        
        # Download
        print("📥 Downloading attachments...")
        wait.until(EC.presence_of_element_located((By.ID, "MailList")))
        email_elem = driver.find_element(By.XPATH, '//*[@id="MailList"]/div/div/div/div/div/div/div/div[2]')
        email_elem.click()
        
        attachments_menu = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//*[@id="focused"]/div[2]/div/div/div/div/div/div/div[2]/button')
        ))
        if attachments_menu:
            attachments_menu[0].click()
            
            download_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[@role="menuitem" and .//span[contains(text(), "Download")]]')
                )
            )
            driver.execute_script("arguments[0].click();", download_button)
            print("✅ Download triggered!")
            time.sleep(5)
    
    except Exception as e:
        print(f"❌ Error: {e}")
        return None
    
    finally:
        driver.quit()
    
    filename = f"Confirm_Payment_{timestamp}.xlsx"
    return wait_and_rename(download_dir, filename)


def process_payment_data():
    """
    Process payment data and filter against summary.
    """
    try:
        paths = get_paths()
    except EnvironmentError as e:
        print(f"\n{e}")
        return False
    
    timestamp = datetime.now().strftime("%Y%m%d")
    
    # Download files
    cf_file = download_emails()
    if not cf_file:
        print("❌ Failed to download payment file")
        return False
    
    # Define columns to keep
    columns_to_keep = [
        "Month", "Name", "New Contract Number", "Old Contract Number", "Product",
        "Payment Amount", "Statement date", "Payment Date", "Channel", "Remark",
        "% Discount", "Portfolio", "Status", "Note", "Responsibility", "OA", "OA-COM", "Address"
    ]
    
    # Read payment file
    sheet_names = pd.ExcelFile(cf_file, engine='openpyxl').sheet_names
    month_sheet_names = [name for name in sheet_names if not re.match(r"(?i)^sheet\d*$", name.strip())]
    
    if not month_sheet_names:
        print("❌ No valid month-named sheets found")
        return False
    
    latest_sheet = month_sheet_names[-1]
    df_temp = pd.read_excel(cf_file, sheet_name=latest_sheet, usecols="B", dtype=str, skiprows=[0])
    last_row = df_temp["Month"].last_valid_index() + 1
    
    df = pd.read_excel(cf_file, sheet_name=latest_sheet, skiprows=[0], usecols=columns_to_keep, nrows=last_row, dtype=str)
    
    # Get summary for deduplication
    summary_dir = paths['summary_report_dir']
    summary_file = None
    try:
        files = glob.glob(os.path.join(summary_dir, "summary_data_file_*.xlsx"))
        if files:
            summary_file = max(files, key=os.path.getmtime)
    except:
        pass
    
    if summary_file:
        # Perform deduplication
        dupplicate_summary = ["pam_code", "EFF_Date", "Pay_Date"]
        summary_df = pd.read_excel(summary_file, usecols=dupplicate_summary, dtype=str)
        summary_df.replace("-", pd.NA, inplace=True)
        summary_df.dropna(inplace=True)
        
        # Clean and merge
        df["New Contract Number"] = df["New Contract Number"].astype(str).str.lstrip("0")
        df["Statement date"] = pd.to_datetime(df["Statement date"]).dt.date
        df["Payment Date"] = pd.to_datetime(df["Payment Date"]).dt.date
        
        summary_df["pam_code"] = summary_df["pam_code"].astype(str).str.lstrip("0")
        summary_df["EFF_Date"] = pd.to_datetime(summary_df["EFF_Date"]).dt.date
        summary_df["Pay_Date"] = pd.to_datetime(summary_df["Pay_Date"]).dt.date
        
        df_renamed = df.rename(columns={
            "New Contract Number": "pam_code",
            "Statement date": "EFF_Date",
            "Payment Date": "Pay_Date"
        })
        
        # Anti-join
        filtered_df = df_renamed.merge(
            summary_df[["pam_code", "EFF_Date", "Pay_Date"]].drop_duplicates(),
            on=["pam_code", "EFF_Date", "Pay_Date"],
            how='left',
            indicator=True
        )
        filtered_df = filtered_df[filtered_df['_merge'] == 'left_only'].drop(columns=['_merge'])
    else:
        filtered_df = df
    
    # Save output
    output_file = os.path.join(paths['output_dir'], f"Confirm_Payment_for_load_{timestamp}.xlsx")
    os.makedirs(paths['output_dir'], exist_ok=True)
    filtered_df.to_excel(output_file, sheet_name="Payment Term History", index=True, index_label="No")
    
    print(f"✅ Output saved: {output_file}")
    print(f"✨ Processed {len(filtered_df)} new records")
    
    return True


def main():
    """
    Main execution function.
    """
    print("\n🚀 Starting payment file processing...\n")
    
    try:
        success = process_payment_data()
        if success:
            print("\n✨ Processing complete!")
            return 0
        else:
            print("\n❌ Processing failed")
            return 1
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
