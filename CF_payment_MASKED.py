"""
CF_payment_MASKED.py - Secure version with masked paths and credentials
================================================================================
⚠️ SECURITY NOTICE:
- All hardcoded credentials have been removed
- All file paths have been masked with [MASKED_*] placeholders
- Use environment variables to provide actual values at runtime

BEFORE DEPLOYING THIS TO PRODUCTION:
1. Set required environment variables (see below)
2. Implement secure credential storage (Azure Key Vault, AWS Secrets Manager)
3. Use OAuth 2.0 for email authentication instead of password
4. Run security scanning tools (bandit, semgrep, etc.)
================================================================================
"""

import os
import sys
from pathlib import Path
from datetime import datetime
import time
import glob
import re

# ⚠️ SECURITY: Import these libraries
try:
    from dotenv import load_dotenv  # For local development only
except ImportError:
    print("⚠️ WARNING: python-dotenv not installed. Using only system environment variables.")

# Load environment variables from .env file (local development ONLY)
# For production, use Azure Key Vault, AWS Secrets Manager, or similar
env_file = Path(__file__).parent / ".env"
if env_file.exists():
    load_dotenv(env_file)

# ============================================================================
# 🔒 SECURE CREDENTIAL HANDLING
# ============================================================================

def get_email_credentials():
    """
    Retrieve email credentials from secure storage.
    
    Priority:
    1. Environment variables (recommended)
    2. Azure Key Vault
    3. AWS Secrets Manager
    
    DO NOT hardcode credentials here!
    """
    EMAIL = os.getenv('EMAIL_ADDRESS')
    PASSWORD = os.getenv('EMAIL_PASSWORD')
    
    if not EMAIL or not PASSWORD:
        raise EnvironmentError(
            "Missing email credentials! Set environment variables:\n"
            "  EMAIL_ADDRESS - Your email address\n"
            "  EMAIL_PASSWORD - Your email password (use app-specific password)\n"
            "Better: Use Microsoft Graph API with OAuth 2.0 instead"
        )
    
    return EMAIL, PASSWORD


def get_file_paths():
    """
    Retrieve file paths from environment variables or configuration.
    
    This prevents hardcoding paths that reveal system architecture.
    """
    paths = {
        'download_dir': os.getenv('DOWNLOAD_DIR', '[MASKED_DOWNLOAD_DIR]'),
        'summary_report_dir': os.getenv('SUMMARY_REPORT_DIR', '[MASKED_SUMMARY_DIR]'),
        'output_dir': os.getenv('OUTPUT_DIR', '[MASKED_OUTPUT_DIR]'),
    }
    
    # Validate paths are not masked (meaning they weren't set)
    for key, value in paths.items():
        if value.startswith('[MASKED'):
            print(f"⚠️ WARNING: {key} not configured!")
    
    return paths


# ============================================================================
# SELENIUM AND WEB AUTOMATION
# ============================================================================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import shutil


def wait_and_rename(download_dir, new_filename, timeout=30):
    """
    Wait for file download to complete and rename it.
    
    Args:
        download_dir: Directory where files are downloaded
        new_filename: New name for the downloaded file
        timeout: Maximum seconds to wait for download
    
    Returns:
        Path to renamed file, or None if download failed
    """
    print("⏳ Waiting for download to finish...")
    end_time = time.time() + timeout
    downloaded_file = None

    while time.time() < end_time:
        files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload")]
        if files:
            # Get the most recently modified file
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


def setup_browser_options(download_dir):
    """
    Configure Chrome browser options with security best practices.
    
    Args:
        download_dir: Directory for downloaded files
    
    Returns:
        Configured ChromeOptions object
    """
    options = Options()
    
    # Suppress verbose logging
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")  # Suppress INFO logs
    
    # Security-focused preferences
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    return options


def download_from_email():
    """
    ⚠️ SECURITY WARNING: This function uses Selenium to automate email login.
    
    BETTER APPROACH: Use Microsoft Graph API with OAuth 2.0
    - Eliminates need for password storage
    - Supports MFA
    - Provides audit logging
    - More secure and reliable
    
    See: https://docs.microsoft.com/en-us/graph/api/
    """
    
    # Get credentials from secure storage
    try:
        EMAIL, PASSWORD = get_email_credentials()
    except EnvironmentError as e:
        print(f"❌ Error: {e}")
        return None
    
    # Get paths from configuration
    paths = get_file_paths()
    download_dir = paths['download_dir']
    
    # Ensure download directory exists
    os.makedirs(download_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d")
    check_load_output_path = os.path.join(download_dir, f"Confirm_Payment_{timestamp}.xlsx")
    
    if Path(check_load_output_path).exists():
        print(f"✅ File already exists: {check_load_output_path}")
        return check_load_output_path
    
    # Setup browser
    options = setup_browser_options(download_dir)
    driver = webdriver.Chrome(service=Service(), options=options)
    
    try:
        # Open Outlook Web
        driver.get("https://outlook.office.com/mail/inbox")
        wait = WebDriverWait(driver, 20)
        
        # Enter email
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]'))).send_keys(EMAIL)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
        time.sleep(3)
        
        # Enter password
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))).send_keys(PASSWORD)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
        time.sleep(3)
        
        # Handle "Stay signed in?" dialog (if appears)
        try:
            wait.until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
        except:
            pass  # Dialog not present
        
        time.sleep(3)
        
        # Search for emails
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchBoxId-Mail"]'))).click()
        
        # Setup search filters
        wait.until(EC.presence_of_element_located((By.ID, 'filtersButtonId')))
        filter_box = driver.find_element(By.ID, 'filtersButtonId')
        filter_box.click()
        time.sleep(1)
        
        # Search for "Confirm Payment" subject
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Subject-ID"]')))
        search_box.click()
        search_box.send_keys('Confirm Payment')
        
        # Select today's date
        current_date = datetime.now().strftime("%m/%d/%Y")
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Select a date"]')))
        date_input.click()
        date_input.send_keys(current_date)
        
        # Click search
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "ms-Button--primary") and @type="button"]')))
        search_button.click()
        
        # Wait for results
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class,"lvHighlightSubjectClass")]')))
        
        # Download attachments
        max_emails = 10
        emails = driver.find_elements(By.XPATH, '//*[@id="MailList"]/div/div/div/div/div/div/div/div')[:max_emails]
        
        print(f"📧 Found {len(emails)} email(s).")
        
        if emails:
            # Click first email
            email = driver.find_element(By.XPATH, '//*[@id="MailList"]/div/div/div/div/div/div/div/div[2]')
            email.click()
            
            # Download attachment
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
                print("✅ Attachment download triggered!")
                time.sleep(5)
    
    except Exception as e:
        print(f"❌ Error during attachment download: {e}")
        return None
    
    finally:
        driver.quit()
    
    # Rename and return downloaded file
    filename = f"Confirm_Payment_{timestamp}.xlsx"
    return wait_and_rename(download_dir, filename)


# ============================================================================
# DATA PROCESSING
# ============================================================================

def get_latest_file(directory, file_pattern):
    """
    Find the latest updated file matching the pattern in a directory.
    
    Args:
        directory: Directory to search
        file_pattern: Glob pattern for files
    
    Returns:
        Path to most recently modified file, or None
    """
    files = glob.glob(os.path.join(directory, file_pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def process_payment_file(cf_file, output_path):
    """
    Process confirm payment file and filter against summary report.
    
    Args:
        cf_file: Path to Confirm Payment Excel file
        output_path: Path for output file
    
    Returns:
        DataFrame with filtered results
    """
    
    # Define columns to keep
    columns_to_keep = [
        "Month", "Name", "New Contract Number", "Old Contract Number", "Product",
        "Payment Amount", "Statement date", "Payment Date", "Channel", "Remark",
        "% Discount", "Portfolio", "Status", "Note", "Responsibility", "OA", "OA-COM", "Address"
    ]
    
    # Read Excel file and find latest month sheet
    sheet_names = pd.ExcelFile(cf_file, engine='openpyxl').sheet_names
    month_sheet_names = [name for name in sheet_names if not re.match(r"(?i)^sheet\d*$", name.strip())]
    
    if not month_sheet_names:
        raise ValueError("No valid month-named sheets found in Excel file.")
    
    latest_sheet = month_sheet_names[-1]
    
    # Read payment data
    df_temp = pd.read_excel(cf_file, sheet_name=latest_sheet, usecols="B", dtype=str, skiprows=[0])
    last_row = df_temp["Month"].last_valid_index() + 1
    
    df = pd.read_excel(
        cf_file,
        sheet_name=latest_sheet,
        skiprows=[0],
        usecols=columns_to_keep,
        nrows=last_row,
        dtype=str
    )
    
    # Get summary report for comparison
    paths = get_file_paths()
    summary_report_dir = paths['summary_report_dir']
    
    summary_file = get_latest_file(summary_report_dir, "summary_data_file_*.xlsx")
    if not summary_file:
        print("⚠️ Warning: Summary report not found. Skipping deduplication.")
        return df
    
    # Read summary for deduplication
    dupplicate_summary = ["pam_code", "EFF_Date", "Pay_Date"]
    dupplicate_payment = ["New Contract Number", "Statement date", "Payment Date"]
    
    summary_df = pd.read_excel(summary_file, usecols=dupplicate_summary, dtype=str)
    summary_df.replace("-", pd.NA, inplace=True)
    summary_df.dropna(inplace=True)
    summary_df.reset_index(drop=True, inplace=True)
    
    # Normalize dates
    df["New Contract Number"] = df["New Contract Number"].astype(str).str.lstrip("0")
    df["Statement date"] = pd.to_datetime(df["Statement date"]).dt.date
    df["Payment Date"] = pd.to_datetime(df["Payment Date"]).dt.date
    
    summary_df["pam_code"] = summary_df["pam_code"].astype(str).str.lstrip("0")
    summary_df["EFF_Date"] = pd.to_datetime(summary_df["EFF_Date"]).dt.date
    summary_df["Pay_Date"] = pd.to_datetime(summary_df["Pay_Date"]).dt.date
    
    # Rename columns for merge
    df_renamed = df.rename(columns={
        "New Contract Number": "pam_code",
        "Statement date": "EFF_Date",
        "Payment Date": "Pay_Date"
    })
    
    # Clean data
    df_renamed.dropna(subset=["pam_code"], inplace=True)
    df_renamed = df_renamed[df_renamed["pam_code"].astype(str).str.strip() != ""]
    df_renamed = df_renamed[df_renamed["pam_code"].astype(str).str.strip().str.isnumeric()]
    
    # Anti-join to find new records
    filtered_df = df_renamed.merge(
        summary_df[dupplicate_summary].drop_duplicates(),
        on=dupplicate_summary,
        how='left',
        indicator=True
    )
    
    # Keep only rows not in summary
    filtered_df = filtered_df[filtered_df['_merge'] == 'left_only'].drop(columns=['_merge'])
    
    return filtered_df


def main():
    """
    Main execution function.
    """
    try:
        timestamp = datetime.now().strftime("%Y%m%d")
        paths = get_file_paths()
        output_dir = paths['output_dir']
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        print("🚀 Starting payment file processing...")
        
        # Download from email
        print("\n📥 Downloading confirm payment file...")
        cf_file = download_from_email()
        
        if not cf_file:
            print("❌ Failed to download payment file.")
            return False
        
        print(f"✅ Downloaded: {cf_file}")
        
        # Process payment file
        print("\n🔄 Processing payment data...")
        filtered_df = process_payment_file(cf_file, output_dir)
        
        # Save output
        output_file = os.path.join(output_dir, f"Confirm_Payment_for_load_{timestamp}.xlsx")
        filtered_df.to_excel(
            output_file,
            sheet_name="Payment Term History",
            index=True,
            index_label="No"
        )
        
        print(f"✅ Output saved: {output_file}")
        print(f"\n✨ Processing complete! Processed {len(filtered_df)} new records.")
        
        return True
    
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
