# Multi-file-excel-automation
Process unlimited Excel files automatically: clean data by processing each, standardize columns, merge files, and generate professional reports with visual tools.
#  Multi-File Excel Processor

## Stop wasting hours on Excel. Let the code do the work.

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Made with Pandas](https://img.shields.io/badge/Made%20with-Pandas-red.svg)](https://pandas.pydata.org)

---

## The Real Story

A client once told me: *"I spend every Friday afternoon opening 12 Excel files, copying data, fixing errors, and creating reports. I hate Fridays."*

That stuck with me.

So I built a tool that turns 3 hours of manual work into 30 seconds of waiting.

**This is that tool.**

---

## What It Does (In Plain English)

You have Excel files. Maybe 5. Maybe 50. Maybe from different departments with different column names and different date formats.

This script:
- **Finds** every Excel file in your folder
- **Cleans** messy data (missing values, duplicates, blank rows)
- **Fixes** column names (whether it says "Amount", "Sale", or "Total", it becomes "Sales")
- **Standardizes** dates (turns "Jan 15, 2024", "15/01/2024", "2024-01-15" into the same format)
- **Combines** everything into one master sheet
- **Summarizes** your data by month (totals, counts, averages)
- **Formats** your report like a pro (blue headers, borders, currency)
- **Shows** a chart of your monthly trends

And if one file is corrupt? The others still work. The error log tells you which file failed.

---

## Before vs After

### Before (Manual Process)
1-Open January.xlsx

2-Copy data

3-Paste into master sheet

4-Repeat for February, March... 9 more times

5-Fix dates (some are DD/MM, some are MM/DD)

6-Fix column names (one file says "Amount", another says "Total")

7-Remove duplicate rows

8-Fill missing values

9-Create pivot table for monthly summary

Format headers, add borders, make it look professional

Double-check everything

Send to manager
Time: 2-3 hours
Frustration: High
Chance of error: Very high

### After (This Tool)
Put all files in a folder
Run: python main.py
Get your report

Time: 30 seconds
Frustration: Zero
Chance of error: Zero

## What You Get

### 1. Master Sheet (All Data Combined)
Every sale from every file, cleaned and standardized.

| Date | Month | Product | Sales | Quantity | Source |
|------|-------|---------|-------|----------|--------|
| 2024-01-15 | January | Laptop | $1,200.00 | 1 | January.xlsx |
| 2024-01-15 | January | Mouse | $25.00 | 2 | January.xlsx |
| 2024-02-01 | February | Laptop | $1,400.00 | 1 | February.xlsx |

### 2. Monthly Summary (What You Actually Care About)
No more pivot tables. Just the numbers you need.

| Month | Total Sales | Orders | Average Sale |
|-------|-------------|--------|--------------|
| January | $15,230.50 | 18 | $846.14 |
| February | $12,450.00 | 15 | $830.00 |
| March | $18,920.75 | 22 | $860.03 |

*Plus a line chart showing your sales trend over months.*

### 3. Error Log (If Something Goes Wrong)
If a file is corrupt or has the wrong format, it's listed here. The rest of your files still process.

| File | Error |
|------|-------|
| corrupt_file.xlsx | Excel file cannot be read |
| empty_file.xlsx | No data found |
---
| Real World Problem | How This Tool Solves It |
|--------------------|-------------------------|
| One file says "Amount", another says "Total" | Both become "Sales" |
| One file says "Qty", another says "Units" | Both become "Quantity" |
| One file says "Item", another says "Prod" | Both become "Product" |
| Dates like "01/15/2024" and "15-01-2024" | All become "2024-01-15" |
| Missing sales amount | Filled with product average |
| Missing quantity | Set to 1 |
| Missing product | Labeled "Unknown" |
| Missing date | Set to first day of month |
| Duplicate rows | Removed automatically |
| Blank rows | Removed automatically |
| "TOTAL" or "GRAND TOTAL" rows | Removed automatically |

## Quick Start (3 Minutes)

### Step 1: Get the code
```bash
git clone https://github.com/MyDAO7/multi-file-excel-processor.git
cd multi-file-excel-processor
**** Step 2: Install requirements
bash
pip install -r requirements.txt
**Step 3: Add your files**
Copy all your Excel files into the client_files folder.

**Step 4: Run the script**
bash
python main.py
**Step 5: Get your report**
Look for Sales_report_Monthly.xlsx in the same folder.
### Project Structure
multi-file-excel-processor/
│
├── main.py                 ← The script (run this)
├── requirements.txt        ← What it needs to work
├── README.md              ← This file
│
├── client_files/          ←  PUT YOUR FILES HERE
│   ├── January.xlsx
│   ├── February.xlsx
│   └── March.xlsx
│
└── Sales_report_Monthly.xlsx   ←  YOUR REPORT
 all requirements are in requirement.txt


Did this save you time?
 Star this repo so others can find it.

Questions? Open an issue on GitHub. I usually respond within a day.
