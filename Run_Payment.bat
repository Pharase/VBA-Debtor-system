@echo off
REM Activate conda environment and run the script

CALL "C:\Users\<user>\miniconda3\Scripts\activate.bat" activate base
python "C:\card_data\processing\program\CF_payment.py"
python "C:\card_data\processing\program\macro_load_file.py"
pause
