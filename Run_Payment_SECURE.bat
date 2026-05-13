@echo off
REM Secure batch file - Uses environment variables instead of hardcoded paths
REM
REM REQUIREMENTS:
REM - .env file must be created with:
REM   - EMAIL_ADDRESS
REM   - EMAIL_PASSWORD (app-specific password)
REM   - DOWNLOAD_DIR
REM   - OUTPUT_DIR
REM   - MACRO_FILE_PATH
REM
REM - Python environment must be set up
REM

echo.
echo ========================================
echo Secure Payment Processing
echo ========================================
echo.

REM Activate conda environment
echo [1/3] Activating Python environment...
CALL "C:\Users\PAMC-NB-Alpha\miniconda3\Scripts\activate.bat" activate base

if errorlevel 1 (
    echo ERROR: Failed to activate Python environment
    pause
    exit /b 1
)

REM Run payment download and processing
echo.
echo [2/3] Running payment file processing...
echo (Uses CF_payment_SECURE.py - credentials from .env)
python "C:\Pam_card\processing\program\CF_payment_SECURE.py"

if errorlevel 1 (
    echo ERROR: Payment processing failed
    pause
    exit /b 1
)

REM Run macro execution
echo.
echo [3/3] Running Excel macro...
echo (Uses macro_load_file_SECURE.py - paths from .env)
python "C:\Pam_card\processing\program\macro_load_file_SECURE.py"

if errorlevel 1 (
    echo WARNING: Macro execution failed, but payment processing completed
)

echo.
echo ========================================
echo Processing Complete!
echo ========================================
echo.
echo NOTE: All sensitive data is now configured via .env file
echo Make sure .env is in .gitignore and never committed
echo.
pause
