@echo off
REM Activate conda environment and run the script

CALL "C:\Users\<user>\miniconda3\Scripts\activate.bat" activate base
python "C:\card_data\processing\program\Summary_transaction_program-v2.py"
pause
