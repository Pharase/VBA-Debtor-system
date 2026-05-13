"""
cutoff_sending_SECURE.py - Secure email sending with credentials from environment
================================================================================
⚠️ SECURITY NOTICE:
- All hardcoded email credentials have been REMOVED
- All hardcoded email addresses are now configurable
- All hardcoded paths are now configurable

REQUIRED ENVIRONMENT VARIABLES:
- EMAIL_ADDRESS: Your email address
- EMAIL_PASSWORD: Your app-specific password (for Office 365)
- TO_EMAIL: Recipient email address
- CC_EMAILS: Comma-separated CC email addresses (optional)

BEFORE DEPLOYING:
1. Create .env file with your credentials
2. Use app-specific password for Office 365 (NOT your main password)
3. RECOMMENDATION: Migrate to Microsoft Graph API with OAuth 2.0
================================================================================
"""

import smtplib
import os
import glob
import sys
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime

try:
    from dotenv import load_dotenv
except ImportError:
    print("⚠️ WARNING: python-dotenv not installed. Using system environment variables only.")

# Load environment variables
env_file = Path(__file__).parent / ".env"
if env_file.exists():
    load_dotenv(env_file)


def get_email_credentials():
    """
    Retrieve email credentials from environment variables.
    
    Returns:
        tuple: (email_address, password)
        
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
            "\n📝 Create .env file with:\n"
            "EMAIL_ADDRESS=your.email@company.com\n"
            "EMAIL_PASSWORD=your-app-specific-password"
        )
    
    return email, password


def get_recipient_emails():
    """
    Retrieve recipient emails from environment variables.
    
    Returns:
        tuple: (to_email, cc_emails_list)
        
    Raises:
        EnvironmentError: If TO_EMAIL not set
    """
    to_email = os.getenv('TO_EMAIL')
    if not to_email:
        raise EnvironmentError(
            "❌ Missing TO_EMAIL environment variable!\n"
            "Set this in your .env file."
        )
    
    cc_emails = os.getenv('CC_EMAILS', '')
    cc_list = [e.strip() for e in cc_emails.split(',') if e.strip()]
    
    return to_email, cc_list


def get_report_paths():
    """
    Retrieve report directories from environment variables.
    
    Returns:
        dict: Report directories and patterns
    """
    report_dirs = {
        "Daily Report": (os.getenv('DAILY_REPORT_DIR', '[MASKED_DAILY_REPORT_DIR]'), "DailyReport_*.xlsx"),
        "Summary Report": (os.getenv('SUMMARY_REPORT_DIR', '[MASKED_SUMMARY_DIR]'), "summary_data_file_*.xlsx"),
        "Summary Report Cut": (os.getenv('SUMMARY_REPORT_DIR', '[MASKED_SUMMARY_DIR]'), "summary_data_file_*-cut.xlsx"),
    }
    
    return report_dirs


def get_latest_file(directory, file_pattern):
    """
    Find the latest updated file matching the pattern.
    """
    try:
        files = glob.glob(os.path.join(directory, file_pattern))
        if not files:
            return None
        return max(files, key=os.path.getmtime)
    except:
        return None


def send_email():
    """
    Send email with report attachments.
    """
    try:
        # Get credentials
        email_address, email_password = get_email_credentials()
        to_email, cc_emails = get_recipient_emails()
        
    except EnvironmentError as e:
        print(f"\n{e}")
        return False
    
    # Email configuration
    SMTP_SERVER = "smtp-mail.outlook.com"
    SMTP_PORT = 587
    current_date = datetime.today().strftime('%d-%m-%Y')
    
    # Email content
    SUBJECT = f"Update Data Daily {current_date}"
    BODY = f"เรียนทีม QMC, Legal\n\nUpdate Data Daily {current_date}\n\nขอบคุณครับ\n\n"
    
    # Get report paths
    report_dirs = get_report_paths()
    attachment_paths = [get_latest_file(dir, pattern) for dir, pattern in report_dirs.values()]
    attachment_paths = [f for f in attachment_paths if f]
    
    # Create email
    msg = EmailMessage()
    msg["From"] = email_address
    msg["To"] = to_email
    if cc_emails:
        msg["Cc"] = ", ".join(cc_emails)
    msg["Subject"] = SUBJECT
    msg.set_content(BODY)
    
    # Attach files
    if attachment_paths:
        for file_path in attachment_paths:
            with open(file_path, "rb") as file:
                file_data = file.read()
                file_name = os.path.basename(file_path)
                msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
                print(f"📎 Attached: {file_name}")
    else:
        print("⚠️ No report files found. Sending email without attachments.")
    
    # Send email
    try:
        print(f"\n📧 Sending email to {to_email}...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure connection
            server.login(email_address, email_password)  # Login
            server.send_message(msg)  # Send email
        print("✅ Email sent successfully!")
        return True
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        return False


def main():
    """
    Main execution function.
    """
    print("\n🚀 Starting email sending process...\n")
    
    try:
        success = send_email()
        if success:
            print("\n✨ Email process complete!")
            return 0
        else:
            print("\n❌ Email process failed")
            return 1
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
