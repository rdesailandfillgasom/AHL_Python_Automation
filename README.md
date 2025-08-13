# AHL_Python_Automation
Arbor Hills Landfill – CSV Appending &amp; Rolling Data Automation

1. Purpose of the Project
This project was built to automate the consolidation of daily landfill monitoring CSV data into a central Excel workbook used for tracking and analysis.
Manual copying and pasting is:
Time-consuming
Error-prone
Inconsistent in formatting

This script solves that by:
Automatically detecting new CSV files in a specified folder.
Cleaning and standardizing the data.
Appending it to the Appended_Data and rolling_data sheets in a master Excel file.
Preventing duplicate entries.
Skipping irrelevant files (like probe data).

2. Core Workflow
The script runs in these stages:
Stage 1 – User Prompt & Excel Shutdown
Why: Excel keeps files locked if they’re open, which can cause write failures.
Logic:
Warns the user that all open Excel windows will be closed.
Asks for confirmation before proceeding.
Uses win32com.client to close all Excel instances.
Stage 2 – Loading the Master Workbook
Checks if the AHL_Rolling_python.xlsx file exists.

Loads:
Appended_Data sheet → holds the complete dataset (all columns + a blank row after each file’s data).
rolling_data sheet → holds only essential monitoring columns for rolling analysis.
If sheets do not exist, initializes empty DataFrames.

Stage 3 – Identifying Eligible CSV Files
Reads all files in the target folder (folder_path).

Filters out:
Files not ending in .csv
Files already processed (tracked in the Source_File column)
Files containing "probe" (case-insensitive) in their name.

This ensures:
No duplicates are added.
Probe-related data is excluded automatically.

Stage 4 – Reading & Cleaning CSV Data
For each eligible CSV:
Load Data using pandas.read_csv.

Data Cleaning:
Convert "NR" or "NA" (case-insensitive, with/without spaces) to None.
Convert ">>>" in the H2S_PPM column to 2001 (threshold exceedance indicator).
Column Alignment:
Reorder and filter columns to match the master sheet’s structure (base_columns).
Add missing columns as blank (None).

Stage 5 – Appending to DataFrames
Add a Source_File column to track which CSV file the data came from.
Append a blank row after the CSV data in Appended_Data.
Create a filtered subset (rolling_filtered) containing only the rolling_columns for rolling_data.

Stage 6 – Writing Back to Excel
Use pandas.ExcelWriter with openpyxl in append mode to update:
Appended_Data
rolling_data

Replace existing sheets with updated versions.

Handle permission errors gracefully (alerts if the file is still open).
