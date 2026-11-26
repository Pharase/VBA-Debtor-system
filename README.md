# 🧾 VBA Debtor System – Automated Payment Processing & Recording Workflow

A hybrid **Excel VBA + Python** solution designed to automate the process of collecting, recording, and summarizing debtor payment transactions.  
The system integrates email attachment processing, data comparison, file generation, and email dispatch into a seamless workflow.

---

## ⚙️ How It Works (Overview)

1. **Start the automation**  
   Copy `.env.template` to `.env`, fill in your values, then run `Run_Payment_SECURE.bat` to kick off the entire process.

2. **Email attachment download & preprocessing**  
   - Automatically saves attached payment files from email.
   - Compares new data with existing summary to extract only **new payment records**.

3. **Trigger payment generation (Excel)**  
   - Opens `payment.xlsm` to generate a file used by the main VBA program.

4. **Transaction recording in UI system**  
   - Launch `trans_program17_4.xlsm`, which includes a VBA-based **User Interface (UI)**.
     - ✅ Press **"Load Payment Info"** to import data.
     - 📝 Press **"Update Value"** to record payment history into each debtor’s card file.
     - 🧾 Press **"Update Datadaily"** to export the latest record to a daily template.
   - File automatically closes once done.

5. **Generate summary & send report**  
   - Run `Run_Summary.bat` to:
     - Generate final summary file.
     - Attach specific files.
     - Send summary via **email**.

---

## 🧩 Components

| File/Script                | Description                                                                 |
|----------------------------|-----------------------------------------------------------------------------|
| `Run_Payment_SECURE.bat`   | Starts the entire workflow (credentials and paths come from `.env`)        |
| `payment.xlsm`             | Prepares intermediate data file                                            |
| `trans_program17_4.xlsm`   | Main UI for recording transactions per debtor                              |
| `summary.py`               | Python script to generate and email summary reports                        |
| `template/datadaily.xlsx`  | Template file updated with latest record                                    |
| `summary_file.xlsx`        | Final output file showing summarized result                                |

---

## 📦 Tech Stack

- **Microsoft Excel (VBA)**
- **Python 3**
- **Win32com & openpyxl** (for file automation and email handling)
- **Windows Batch Script**

---

## 🛠 Requirements

- Windows OS with Microsoft Excel installed
- Python 3 environment
- Required Python packages:  
  ```bash
  pip install openpyxl pandas pywin32

---

## 🔐 Security notes

- Credentials and file paths are read from `.env` (see `.env.template`); `.env` is git-ignored.
- The real workbooks contain debtor data and are **not** in this repository. Their VBA code is exported
  to [`vba/`](vba/) so it can be read and reviewed; `_*.xlsx` files are header-only templates.
