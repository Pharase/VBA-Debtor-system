@echo off
REM Activate conda environment and run the script

CALL "%USERPROFILE%\miniconda3\Scripts\activate.bat" base
cd /d "%~dp0"
python "%~dp0Summary_transaction_program-v2.py"
pause
