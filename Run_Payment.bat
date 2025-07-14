@echo off
REM Activate conda environment and run the script

CALL "C:\Users\PAMC-NB-Alpha\miniconda3\Scripts\activate.bat" activate base
python "C:\Pam_card\processing\program\CF_payment.py"
python "C:\Pam_card\processing\program\macro_load_file.py"
pause
