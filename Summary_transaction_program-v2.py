#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import win32gui
import pandas as pd
import xlwings as xw
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime


# Function to convert column index to Excel column letter
def column_index_to_letter(index):
    # Convert zero-indexed column index to Excel column letter
    result = ""
    while index >= 0:
        result = chr(index % 26 + 65) + result
        index = index // 26 - 1
    return result
    
# Define a function to copy and paste data to a specific Excel file
def copy_and_paste_data(file_path, data_to_copy):
    try:
        # Open the Excel file
        wb = xw.Book()
        ws = wb.sheets['Sheet1']  # Change to the target sheet name

        # Find the next available row in the target sheet
        next_row = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

        # Convert all data to strings to preserve leading zeros
        data_to_copy_str = [[str(cell) if cell is not None else '' for cell in row] for row in data_to_copy]

        # Set number format to text ('@') for the entire range where data will be pasted
        last_column = 'A' + column_index_to_letter(len(data_to_copy_str[0]) - 1)
        ws.range(f'A{next_row}:{last_column}{next_row + len(data_to_copy_str)}').number_format = '@'

        # Paste the data into the target sheet
        ws.range(f'A{next_row}').value = data_to_copy_str

        # Close the workbook
        wb.save(file_path)
        wb.close()

    except Exception as e:
        print(f"Error copying and pasting data to file {file_path}: {e}")

# Define a function to process each Excel file
def process_excel_file(file_path):
    try:
        # Open the Excel file
        wb = xw.Book(file_path)
        ws = wb.sheets['Payment Term History']
        
        # Read the value from cell K3
        cell_count = ws.range(3, 11).value
        if cell_count is None:
            print(f"Cell K3 in file {file_path} is empty.")
            return None

        # Convert cell_value to an integer
        cell_count = int(cell_count)

        # Initialize a list to store cell values
        cell_values = []

        for i in range(7, 7 + cell_count):
            # Construct the range representing the entire row from column A to EG
            row_range = ws.range(f'A{i}:EG{i}').value

            if row_range[0] is not None and isinstance(row_range[0], (int, float)) or (isinstance(row_range[0], str) and row_range[0].isdigit()):
                try:
                    temp_value = str(int(row_range[0]))
                    if len(temp_value) == 7:
                        row_range[0] = '0' + temp_value
                except ValueError:
                    print(f"Conversion error for cell A{i}: {row_range[0]}")
                    
            cell_values.append(row_range)

        # Close the workbook
        wb.close()

        return cell_values

    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

def process_excel_file_pandas(file_path):
    try:
        # Read only the necessary cell to get the row count from I3
        cell_value_df = pd.read_excel(file_path, sheet_name='Payment Term History', 
                                      usecols="K", skiprows=2, nrows=1, header=None)
        cell_count = cell_value_df.iloc[0, 0]
        
        if pd.isna(cell_count) or cell_count == "":
            print(f"Cell K3 in file {file_path} is empty.")
            return None
        
        try:
            cell_count = int(cell_count)
        except ValueError:
            print(f"Value in cell K3 of file {file_path} is not a valid integer.")
            return None
        
        # Read the specified number of rows
        rows_df = pd.read_excel(file_path, sheet_name='Payment Term History', 
                                skiprows=6, nrows=cell_count, 
                                usecols="A:EG")
        
        # Option 1: Replace empty strings with a fixed value (e.g., 0, None)
        rows_df.replace("", 0, inplace=True)
        
        # Option 2: Drop rows where all elements are empty strings
        # rows_df.dropna(how='all', inplace=True)
        
        # Option 3: Drop rows where any element is an empty string
        # rows_df.dropna(inplace=True)
        
        # Convert DataFrame to a list of rows
        cell_values = rows_df.values.tolist()

        return cell_values

    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

# Define the base directory containing the Excel files
base_folder = r"C:\Pam_card\summary"
subfolders = ["done-add-date"]  # You can add more subfolders here "g1", "g2", "g3", "g4", 
main_data = r"C:\Pam_card\_Data_CutOff_v19.xlsx"

# Get today's date
today = datetime.today()

# Format the date as yymmdd
formatted_date = today.strftime('%y%m%d')

# Open the main Excel file
main_file = xw.Book(main_data, update_links=True)

# Get Excel application from window handle
xl_app = xw.apps.active

# Drop specific columns by their indices or column letters
columns_to_drop = ['DF', 'CP', 'BA', 'AZ', 'AY', 'AX', 'AV', 'AJ', 'AI', 'A']

# Iterate over each subfolder
for subfolder in subfolders:
    # Define the full folder path
    folder_path = os.path.join(base_folder, subfolder)
    
    # Define the output file path for the combined data
    temp_file = os.path.join(r"C:\Pam_card\summary", f"summary_data_{formatted_date}_temp.xlsx")

    # Collect file paths for Excel files in the current subfolder
    file_list = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path) if filename.startswith("StatementCard_") and filename.endswith(".xlsx")]
    total_files = len(file_list)
    dest_row = 1
    
    # Initialize a list to store data from all files
    all_data = []

    # Process files using multithreading
    with ThreadPoolExecutor(max_workers=1) as executor:
        futures = [executor.submit(process_excel_file, file_path) for file_path in file_list]
        
        # Add a progress bar using tqdm
        for future in tqdm(futures, total=len(file_list), desc=f"Processing {subfolder}"):
            data = future.result()
            if data:
                all_data.extend(data)

    # Copy and paste data to the output file
    copy_and_paste_data(temp_file, all_data)

wb = xw.Book(temp_file)
ws = wb.sheets['Sheet1'] 
for col_index in columns_to_drop:
    ws.range(ws.cells(1, col_index), ws.cells(ws.cells.last_cell.row, col_index)).api.Delete()

print("Data extraction and combination completed.")

summary_temp = r"C:\Pam_card\summary\_summary_template_v2.xlsx"
destb = xw.Book(summary_temp)
dests = destb.sheets['transaction_record']

dests.range('ED17').value = main_data

# Find the next available row in copy  sheet
next_row = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row

cell_values = []
count_change = 0

for i in range(1, next_row + 1):
# Construct the range representing the entire row from column A to CV
    row_range = ws.range(f'A{i}:CV{i}').value
    if len(str(row_range[0])) == 8:
        temp_value = str(int(row_range[0]))
        row_range[0] = '0' + temp_value
        count_change = count_change + 1
    cell_values.append(row_range)

# Paste the data into the target sheet
count_change = count_change + 1
dests.range(f'A2:A{count_change}').number_format = '@'
dests.range(f'A2').value = cell_values

cell_values = []

for i in range(1, next_row + 1):
# Construct the range representing the entire row from column CY to DW
    row_range = ws.range(f'CY{i}:DW{i}').value
    if len(str(row_range[0])) == 8:
        temp_value = str(int(row_range[0]))
        row_range[0] = '0' + temp_value
    cell_values.append(row_range)

# Paste the data into the target sheet
dests.range(f'CY2:CY{count_change}').number_format = '@'
dests.range(f'CY2').value = cell_values

output_file = os.path.join(r"C:\Pam_card\summary", f"summary_data_file_{formatted_date}.xlsx")

# Update Cost
dests = destb.sheets['Cost']
Cost_last_row = int(dests.range('G2').value) 

dests.range(f'E{2}').formula2 = (
        '=IFERROR('
            'MIN(FILTER(transaction_record!DP:DP, ((transaction_record!A:A = [@[pam_code]])*(transaction_record!M:M <> "sold" )))), '
            '[@[bid_price_excluding_vat]]'
        ')'
    )

values = dests.range(f'E2:E{Cost_last_row}').value
if isinstance(values, list) and all(not isinstance(v, list) for v in values):
    values = [[v] for v in values]
dests.range(f'E2:E{Cost_last_row}').value = values

# Refresh all data connections
dests.book.sheets["pivot (3) by period"].api.PivotTables("PivotTable1").PivotCache().Refresh()
dests.book.sheets["pivot (4) by case"].api.PivotTables("PivotTable1").PivotCache().Refresh()
        
# Save and close the workbook 
destb.save(output_file)
destb.close()
wb.save()
wb.close()
main_file.close()

print("Data Summary completed.")

#Cut column
from cut_column import Cut_column_summary
Cut_column_summary(output_file)

#Copy file
from move_file import copy_files_created_today

# Example usage
source_directory1 = "C:/Pam_card/summary/done-add-date"  # Replace with your source directory
target_directory1 = "Z:/CutOff/3.Statement Card"  # Replace with your target directory
copy_files_created_today(source_directory1, target_directory1)

source_directory2 = "C:/Pam_card/summary"  # Replace with your source directory
target_directory2 = "Z:/CutOff/6.Summary"  # Replace with your target directory
copy_files_created_today(source_directory2, target_directory2)

source_directory3 = "Z:/CutOff/4.Daily Report"  # Replace with your source directory
target_directory3 = "C:/Users/PAMC-NB-Alpha/Hylife Group/PAM-Data - Documents/Data For QMC/Payment Term/3-Daily Database"  # Replace with your target directory
copy_files_created_today(source_directory3, target_directory3)

source_directory4 = "C:/Pam_card/summary"  # Replace with your source directory
target_directory4 = "C:/Users/PAMC-NB-Alpha/Hylife Group/PAM-Data - Documents/Data For QMC/Payment Term/7-Summary"  # Replace with your target directory
copy_files_created_today(source_directory4, target_directory4)

source_directory5 = "C:/Pam_card/summary/done-add-date"  # Replace with your source directory
target_directory5 = "C:/Users/PAMC-NB-Alpha/Hylife Group/PAM-Data - Documents/Data For QMC/Payment Term/2-Cut-off Table"  # Replace with your target directory
copy_files_created_today(source_directory5, target_directory5)

#Mailing
from cutoff_sending import send_email
send_email()