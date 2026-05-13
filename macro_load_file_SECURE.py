"""
macro_load_file_SECURE.py - Secure version with hardcoded paths removed
================================================================================
⚠️ SECURITY NOTICE:
- All hardcoded file paths have been removed and replaced with environment variables
- All sensitive data (paths, usernames) are now configurable

REQUIRED ENVIRONMENT VARIABLES:
- MACRO_FILE_PATH: Full path to the macro-enabled Excel file

BEFORE RUNNING:
1. Set environment variables or create .env file
2. Update config.ini with your actual file paths
================================================================================
"""

import win32com.client as win32
import os
import time
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    pass

# Load environment variables
env_file = Path(__file__).parent / ".env"
if env_file.exists():
    load_dotenv(env_file)


def get_macro_file_path():
    """
    Retrieve macro file path from environment variable.
    
    Returns:
        str: Path to macro-enabled workbook
        
    Raises:
        EnvironmentError: If MACRO_FILE_PATH is not set
    """
    file_path = os.getenv('MACRO_FILE_PATH')
    if not file_path:
        raise EnvironmentError(
            "❌ MACRO_FILE_PATH environment variable not set!\n"
            "Set this to your macro-enabled Excel file path."
        )
    return file_path


def run_excel_macro(file_path, macro_name):
    """
    Execute Excel macro from Python with error handling.
    
    Args:
        file_path (str): Full path to .xlsm file
        macro_name (str): Name of macro to execute
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return False

    try:
        # Launch Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False  # Set to True if you want to see Excel
        excel.ScreenUpdating = False  # Improve performance

        # Record current open workbooks before opening the target
        existing_wbs = [wb.Name for wb in excel.Workbooks]

        # Open the workbook
        wb_main = excel.Workbooks.Open(file_path)
        print(f"✅ Opened: {os.path.basename(file_path)}")

        # Run the linked files opener macro first (if exists)
        try:
            print("🔗 Running OpenLinkedFiles macro...")
            excel.Application.Run(f"'{os.path.basename(file_path)}'!OpenLinkedFiles")
            time.sleep(2)
        except Exception as e:
            print(f"⚠️ OpenLinkedFiles macro not found or failed: {e}")

        # Run the target macro
        try:
            print(f"▶️  Running {macro_name} macro...")
            excel.Application.Run(f"'{os.path.basename(file_path)}'!{macro_name}")
            print(f"✅ {macro_name} completed successfully")
        except Exception as e:
            print(f"❌ Error running macro {macro_name}: {e}")
            return False
        
        # Optional wait (if macro takes time to finish opening other files)
        time.sleep(2)

        # Close all newly opened workbooks except the original one
        for wb in excel.Workbooks:
            if wb.Name not in existing_wbs:
                print(f"📁 Closing: {wb.Name}")
                wb.Close(SaveChanges=True)

        # Close the main workbook
        print(f"📁 Closing main workbook...")
        wb_main.Close(SaveChanges=True)

        # Quit Excel
        excel.Quit()
        print("✅ Excel closed successfully")
        
        return True

    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def main():
    """
    Main execution function.
    """
    try:
        # Get macro file path from environment
        macro_file_path = get_macro_file_path()
        macro_to_run = "Create_report_payment"
        
        print(f"\n🚀 Starting macro execution...")
        print(f"📂 Macro file: {os.path.basename(macro_file_path)}")
        print(f"📌 Macro name: {macro_to_run}\n")
        
        success = run_excel_macro(macro_file_path, macro_to_run)
        
        if success:
            print("\n✨ Macro execution completed successfully!")
            return 0
        else:
            print("\n❌ Macro execution failed.")
            return 1
    
    except EnvironmentError as e:
        print(f"\n{e}")
        print("\n📝 Example .env file:")
        print("MACRO_FILE_PATH=C:/path/to/your/_Payment_report_template_v7.xlsm")
        return 1
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
