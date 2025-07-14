import win32com.client as win32
import os
import time

def run_excel_macro(file_path, macro_name):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        # Launch Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False  # Set to True if you want to see Excel

        # Record current open workbooks before opening the target
        existing_wbs = [wb.Name for wb in excel.Workbooks]

        # Open the workbook
        wb_main = excel.Workbooks.Open(file_path)

        # Run the macro
        excel.Application.Run(f"'{os.path.basename(file_path)}'!{"OpenLinkedFiles"}")
        excel.Application.Run(f"'{os.path.basename(file_path)}'!{macro_name}")
        
        # Optional wait (if macro takes time to finish opening other files)
        time.sleep(2)

        # Close all newly opened workbooks except the original one
        for wb in excel.Workbooks:
            if wb.Name not in existing_wbs:
                wb.Close(SaveChanges=True)

        # Close the main workbook
        wb_main.Close(SaveChanges=True)

        # Quit Excel
        excel.Quit()

        print("Macro executed and all files closed.")

    except Exception as e:
        print(f"Error: {e}")

# === Example Usage ===
xlsm_path = r"C:\Pam_card\processing\program\_Payment_report_template_v7.xlsm"
run_excel_macro(xlsm_path,macro_name="Create_report_payment")