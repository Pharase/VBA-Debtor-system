@echo off
REM Activate conda environment and run the script

CALL "C:\Users\PAMC-NB-Alpha\miniconda3\Scripts\activate.bat" activate base
python "C:\Pam_card\processing\program\Summary_transaction_program-v2.py"
pause
