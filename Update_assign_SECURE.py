"""
Update_assign_SECURE.py - Secure version with hardcoded password removed
================================================================================
⚠️ SECURITY NOTICE:
- Hardcoded Excel password 'pam2025' has been REMOVED
- Password now comes from environment variable or secure prompt
- File paths are now configurable

REQUIRED ENVIRONMENT VARIABLES:
- EXCEL_FILE_PASSWORD: Password for encrypted Excel file

BEFORE RUNNING:
1. Set EXCEL_FILE_PASSWORD environment variable
2. Or enter password when prompted
3. Do NOT commit passwords to version control
================================================================================
"""

import pandas as pd
from tkinter import Tk, filedialog, messagebox
import msoffcrypto
import openpyxl
from io import BytesIO
import win32com.client as win32
import os
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


def select_file(title):
    """
    Open file dialog to select Excel file.
    """
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )


def get_excel_password():
    """
    Retrieve Excel password from environment or prompt user.
    
    Returns:
        str: Excel file password
        
    Raises:
        ValueError: If password not provided
    """
    # Try environment variable first
    password = os.getenv('EXCEL_FILE_PASSWORD')
    
    if password:
        print("✅ Password retrieved from environment variable")
        return password
    
    # Prompt user
    root = Tk()
    root.withdraw()
    
    from tkinter.simpledialog import askstring
    password = askstring(
        "Password Required",
        "Enter Excel file password:",
        show="*"
    )
    root.destroy()
    
    if not password:
        raise ValueError("No password provided")
    
    return password


def main():
    """
    Main execution function.
    """
    root = Tk()
    root.withdraw()
    
    try:
        # Select files
        file1_path = select_file("Select the Data Daily File (to be updated)")
        if not file1_path:
            messagebox.showerror("Error", "No daily data file selected.")
            return False
        
        file2_path = select_file("Select the Update Data File")
        if not file2_path:
            messagebox.showerror("Error", "No update data file selected.")
            return False
        
        # Get password
        try:
            password = get_excel_password()
        except ValueError:
            messagebox.showerror("Error", "No password provided.")
            return False
        
        print(f"\n🔐 Decrypting file: {os.path.basename(file1_path)}...")
        
        # Decrypt file
        decrypted = BytesIO()
        with open(file1_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            try:
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to decrypt file: {e}")
                return False
        
        print("✅ File decrypted successfully")
        
        try:
            # Load both files
            print("\n📂 Loading update data...")
            df_update = pd.read_excel(file2_path, sheet_name="ALL")
            update_dict = df_update.set_index("เลขที่สัญญาใหม่").to_dict("index")
            
            # Load workbook with formulas
            print("📂 Loading workbook...")
            wb = openpyxl.load_workbook(decrypted)
            ws = wb["Daily_Report"]
            
            # Define mappings
            mapping = {
                "product": "สินเชื่อ",
                "QMC_status": "Current Status",
                "responsibility": "Owner",
                "oa": "OA",
                "assign_status": "Assign Status",
                "note": "Assign Note",
                "assign_date": "Assign Date"
            }
            
            # Get column indexes
            headers = [cell.value for cell in ws[1]]
            col_indexes = {name: headers.index(name) + 1 for name in mapping}
            id_col_index = headers.index("pam_code") + 1
            
            # Update rows
            print("🔄 Updating records...")
            update_count = 0
            for row in ws.iter_rows(min_row=2):
                pamcode = row[id_col_index - 1].value
                if pamcode in update_dict:
                    update_values = update_dict[pamcode]
                    for col_name, update_col_name in mapping.items():
                        col_idx = col_indexes[col_name]
                        new_value = update_values.get(update_col_name, "")
                        row[col_idx - 1].value = new_value
                    update_count += 1
            
            print(f"✅ Updated {update_count} records")
            
            # Save temporarily
            print("\n💾 Saving changes...")
            wb.save("updated_temp.xlsx")
            
            # Re-encrypt with original password
            print("🔐 Re-encrypting file...")
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb_excel = excel.Workbooks.Open(os.path.abspath("updated_temp.xlsx"))
            wb_excel.SaveAs(os.path.abspath(file1_path), Password=password)
            wb_excel.Close()
            excel.Quit()
            
            # Clean up
            os.remove("updated_temp.xlsx")
            
            print(f"✅ File updated and saved: {file1_path}")
            messagebox.showinfo("Success", f"Updated {update_count} records successfully!")
            
            return True
        
        except Exception as e:
            messagebox.showerror("Error", str(e))
            import traceback
            traceback.print_exc()
            return False
    
    finally:
        root.destroy()


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
