import os
import shutil
from datetime import datetime, date

def copy_files_created_today(source_dir, target_dir):
    # Get today's date
    today = date.today()

    # Check if the source directory exists
    if not os.path.exists(source_dir):
        print(f"Source directory '{source_dir}' does not exist.")
        return

    # Create the target directory if it doesn't exist
    os.makedirs(target_dir, exist_ok=True)

    # Loop through all files in the source directory
    for file_name in os.listdir(source_dir):
        source_file = os.path.join(source_dir, file_name)

        # Skip directories, process files only
        if os.path.isfile(source_file):
            # Get file modification time
            modification_time = datetime.fromtimestamp(os.path.getmtime(source_file))

            # Check if the file was created today
            if modification_time.date() == today:
                target_file = os.path.join(target_dir, file_name)
                shutil.copy(source_file, target_file)
                print(f"Copied: {source_file} -> {target_file}")