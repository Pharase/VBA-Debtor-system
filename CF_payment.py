from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import datetime
import pandas as pd
import re
import time
import os
import glob
import shutil
import sys
from pathlib import Path
from datetime import datetime


def wait_and_rename(download_dir, new_filename, timeout=30):
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
 
timestamp = datetime.now().strftime("%Y%m%d")
download_dir = r"C:\Pam_card\payment\raw_email_attached"
check_load_output_path = os.path.join(download_dir, f"Confirm_Payment_{timestamp}.xlsx")

if Path(check_load_output_path).exists():
    print(f"✅ File already exists: {check_load_output_path}")
else:    
    max_emails = 10
    EMAIL = "paramut.c@pinnacle-amc.co.th"
    PASSWORD = "L@liY220941"
    options = Options()
    
    # Suppress GCM and DevTools logs
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")  # Suppress INFO logs

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=Service(),options=options)

    # Open Outlook Web
    driver.get("https://outlook.office.com/mail/inbox")

    wait = WebDriverWait(driver, 20)

    # Step 2: Enter email
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]'))).send_keys(EMAIL)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
    time.sleep(3)
    # Enter password
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))).send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()
    time.sleep(3)
    # Step 4: Stay signed in? -> Click No (if appears)
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
    except:
        pass  # No "stay signed in" screen

    time.sleep(3)
    # Wait for inbox to load
    driver.find_element(By.XPATH, '//*[@id="searchBoxId-Mail"]').click()
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchBoxId-Mail"]'))).click()
    
    # === SEARCH EMAIL ===
    wait = WebDriverWait(driver, 10)

    # Wait and click the filter button
    wait.until(EC.presence_of_element_located((By.ID, 'filtersButtonId')))
    filter_box = driver.find_element(By.ID, 'filtersButtonId')
    filter_box.click()
    time.sleep(1)

    # Wait and interact with the subject input box
    search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Subject-ID"]')))
    search_box.click()
    search_box.send_keys('Confirm Payment')

    # Select date
    current_date = datetime.now().strftime("%m/%d/%Y")
    date_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Select a date"]')))
    date_input.click()
    date_input.send_keys(current_date)

    # Click search
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "ms-Button--primary") and @type="button"]')))
    search_button.click()

    # Optional: wait for results to load (adjust timeout if needed)
    wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class,"lvHighlightSubjectClass")]')))
    
    wait = WebDriverWait(driver, 10)

    try:
        # ✅ Wait for the mail list to load
        wait.until(EC.presence_of_element_located((By.ID, "MailList")))
        emails = driver.find_elements(By.XPATH, '//*[@id="MailList"]/div/div/div/div/div/div/div/div')[:max_emails]

        print(f"📧 Found {len(emails)} email(s).")
        if not emails:
            print("✅ No emails found. Exiting.")
            #sys.exit()

        # ✅ Click the first email
        email = driver.find_element(By.XPATH, '//*[@id="MailList"]/div/div/div/div/div/div/div/div[2]')
        email.click()

        # Step 1: Click the attachment "..." button (More actions)
        attachments_menu = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//*[@id="focused"]/div[2]/div/div/div/div/div/div/div[2]/button')
        ))
        if attachments_menu:
            attachments_menu[0].click()
        else:
            raise Exception("❌ Could not find the attachment menu button")

        # Step 2: Wait for the pop-up and click the "Download" option
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[@role="menuitem" and .//span[contains(text(), "Download")]]')
            )
        )
        driver.execute_script("arguments[0].click();", download_button)

        print("✅ Attachment download triggered!")
        time.sleep(5)  # Allow time for the file to download

    except Exception as e:
        print("❌ Error during attachment download:", e)

    print("✅ Done downloading attachments.")

    #driver.quit()

filename = f"Confirm_Payment_{timestamp}.xlsx"
CF_file = wait_and_rename(download_dir, filename)
if CF_file == None:
    CF_file = check_load_output_path

columns_to_keep = [
    "Month", "Name", "New Contract Number", "Old Contract Number", "Product",
    "Payment Amount", "Statement date", "Payment Date", "Channel", "Remark",
    "% Discount", "Portfolio", "Status", "Note", "Responsibility", "OA", "OA-COM", "Address"
]

sheet_names = pd.ExcelFile(CF_file, engine='openpyxl').sheet_names 
# Filter out default sheet names like "Sheet1", "Sheet2"
month_sheet_names = [name for name in sheet_names if not re.match(r"(?i)^sheet\d*$", name.strip())]

# Use the last one (you may sort them if needed)
if month_sheet_names:
    latest_sheet = month_sheet_names[-1]
else:
    raise ValueError("No valid month-named sheets found.")

df_temp = pd.read_excel(CF_file, sheet_name=latest_sheet, usecols="B", dtype=str, skiprows=[0])
last_row = df_temp["Month"].last_valid_index() + 1  # +1 because iloc is exclusive

df = pd.read_excel(CF_file,sheet_name=latest_sheet,skiprows=[0],usecols=columns_to_keep, nrows=last_row, dtype=str)

# Function to find the latest updated file in a directory
def get_latest_file(directory, file_pattern):
    files = glob.glob(os.path.join(directory, file_pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)  # Get the most recently modified file
    
REPORT_DIRS = {
        "Summary Report": (r"Z:\CutOff\6.Summary", "summary_data_file_*.xlsx")
    }

# Get the latest files
ATTACHMENT_PATHS = [get_latest_file(dir, pattern) for dir, pattern in REPORT_DIRS.values()]
ATTACHMENT_PATHS = [file for file in ATTACHMENT_PATHS if file]  # Remove None values

# Attach Files
if ATTACHMENT_PATHS:
    for file_path in ATTACHMENT_PATHS:
        with open(file_path, "rb") as file:
            file_data = file.read()
            file_name = os.path.basename(file_path)
        print(f"Check Summary with: {file_name}")
else:
    print("No valid Summary report files found.")

dupplicate_summary = ["pam_code", "EFF_Date", "Pay_Date"]
dupplicate_payment = ["New Contract Number", "Statement date", "Payment Date"]
summary_df = pd.read_excel(file_path, usecols=dupplicate_summary, dtype=str)

# Replace "-" with NaN
summary_df.replace("-", pd.NA, inplace=True)
summary_df.dropna(inplace=True)
summary_df.reset_index(drop=True, inplace=True)

# Strip time part by converting to date
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

# Drop rows where pam_code is NA or empty string
df_renamed.dropna(subset=["pam_code"], inplace=True)
df_renamed = df_renamed[df_renamed["pam_code"].astype(str).str.strip() != ""]
df_renamed = df_renamed[
    df_renamed["pam_code"]
    .astype(str)
    .str.strip()
    .str.isnumeric()
]

# Perform a left-anti join to keep only rows not in summary_df
filtered_df = df_renamed.merge(
    summary_df[dupplicate_summary].drop_duplicates(),
    on=dupplicate_summary,
    how='left',
    indicator=True
)

# Keep only rows that do not match (i.e., '_merge' == 'left_only')
filtered_df = filtered_df[filtered_df['_merge'] == 'left_only'].drop(columns=['_merge'])
filtered_df.to_excel(f"C:/Pam_card/payment/raw_email_attached/Confirm_Payment_for_load_{timestamp}.xlsx",sheet_name="Payment Term History",index=True,index_label="No")

print(f"CF_payment file done! :Confirm_Payment_for_load_{timestamp}.xlsx")