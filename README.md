 📊Eccho beam too (v2)

CDR Desktop Analyzer v2 is a desktop tool for analyzing Call Detail Records (CDRs) from CSV files.
It generates a multi-sheet Excel report (9 tabs) with pre-sorted, styled, and color-coded sheets to make investigation and pattern analysis faster and more intuitive.

🚀 Key Features
🔹 Core Functionality

Multi-file CDR Import – Import multiple CSV files in one session

Automatic Column Handling – Supports varied CDR formats & naming conventions

9 Analysis Sheets – Focused reports for communication, movement, devices, states, and behavior

Pre-sorted Data – Sheets are automatically sorted by most useful metrics (events, duration, dates)

Smart Styling – Headers, important columns, alternating rows, and tab colors for improved UX

Freeze Panes & Autofit – Easy navigation in large datasets

Real-time Progress Updates – Shows progress percentage while processing

Cancel Support – Stop processing mid-way if needed

Detailed Logging – Built-in error and event logging for debugging

📑 Generated Excel Sheets

The analyzer currently produces 9 structured and styled sheets:

_01_CDR_Format – Cleaned & standardized CDR format

_02_Relationship_Call_Frequ – Communication frequency with contacts (sorted by total events & duration)

_03_Cell_ID_Frequency – Cell tower usage frequency (sorted by total events)

_04_Movement_Analysis – Chronological movement timeline (sorted by date/time)

_05_Imei_Used – Device IMEI usage statistics (sorted by total events)

_06_State_Connection – State-wise communication patterns (sorted by total events)

_07_ISD_Call – International calls analysis (sorted by date/time)

_08_Night_Call – Night-time communications (sorted by total events)

_09_Mobile_SwitchOFF – Device switch-off gaps (sorted by start date)

Each sheet has:
✔️ Colored tab for quick identification
✔️ Highlighted important headers & columns
✔️ Alternating row shading for readability

🖥️ User Interface

GUI (tkinter) – Simple, investigator-friendly interface

File Management – Add, preview, validate CSVs before analysis

Batch Processing – Multiple CDRs handled at once

One-click Excel Export – Saves all 9 reports into a single .xlsx file

⚙️ Installation
Requirements

Python 3.7+

Required packages:

pip install pandas numpy openpyxl

Setup

Clone or download the repository

Install the dependencies (above)

Run the application:

python main.py

🛡️ Notes

Optimized for law enforcement, forensic analysis, and telecom investigations

Handles NaT/NaN dates safely in relationship & switch-off reports

Backward compatible with generate_excel_file() used by older modules