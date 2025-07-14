import smtplib
import os
import glob
from email.message import EmailMessage
from datetime import datetime

# Function to find the latest updated file in a directory
def get_latest_file(directory, file_pattern):
    files = glob.glob(os.path.join(directory, file_pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)  # Get the most recently modified file

def send_email():
    current_date = datetime.today().strftime('%d-%m-%Y')

    # Outlook SMTP Server Details
    SMTP_SERVER = "smtp-mail.outlook.com"
    SMTP_PORT = 587

    # Sender Credentials (Use environment variables for security)
    EMAIL_ADDRESS = os.getenv("OUTLOOK_EMAIL", "paramut.c@pinnacle-amc.co.th")
    EMAIL_PASSWORD = os.getenv("OUTLOOK_PASSWORD", "L@liY220941")

    # Recipient Email
    TO_EMAIL = "cut-off-payment@hylife.co.th"
    CC_EMAIL = ["thittaya.y@pinnacle-amc.co.th", "weerachai.t@pinnacle-amc.co.th", "kamonwan@hylife.co.th"]

    # Email Subject & Body
    SUBJECT = f"Update Data Daily {current_date}"
    BODY = f"เรียนทีม QMC, Legal\n\nUpdate Data Daily {current_date} \n\nขอบคุณครับ\n\n"

    # Report Directories
    REPORT_DIRS = {
        "Daily Report": (r"Z:\CutOff\4.Daily Report", "DailyReport_*.xlsx"),
        "Summary Report": (r"Z:\CutOff\6.Summary", "summary_data_file_*.xlsx"),
        "Summary Report Cut": (r"Z:\CutOff\6.Summary", "summary_data_file_*-cut.xlsx"), 
    }

    # Get the latest files
    ATTACHMENT_PATHS = [get_latest_file(dir, pattern) for dir, pattern in REPORT_DIRS.values()]
    ATTACHMENT_PATHS = [file for file in ATTACHMENT_PATHS if file]  # Remove None values

    # Create Email
    msg = EmailMessage()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = TO_EMAIL
    msg["Cc"] = ", ".join(CC_EMAIL)
    msg["Subject"] = SUBJECT
    msg.set_content(BODY)

    # Attach Files
    if ATTACHMENT_PATHS:
        for file_path in ATTACHMENT_PATHS:
            with open(file_path, "rb") as file:
                file_data = file.read()
                file_name = os.path.basename(file_path)
                msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
            print(f"Attached: {file_name}")
    else:
        print("No valid report files found. Sending email without attachments.")

    # Send Email via Outlook SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure connection
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # Login
            server.send_message(msg)  # Send email
        print("✅ Email sent successfully!")
    except Exception as e:
        print(f"❌ Error sending email: {e}")
