# Excel Sheet Segregator (Fast)

Split one Excel sheet into multiple sheets by any column value.

## Table of Contents

1. [What it solves](#what-it-solves)
2. [How to use](#how-to-use)
3. [Key features](#key-features)
4. [Screenshots](#screenshots)
5. [Install](#install)

---

## What it solves

When a large Excel sheet contains mixed data for regions, vendors, teams, or categories, separating it manually is slow and repetitive.

This tool does it in one click.

It takes one selected sheet, groups rows by the column you choose, and creates a new workbook with one sheet per group.

---

## How to use

1. Open the app.
2. Browse and select your Excel file.
3. Choose the sheet.
4. Choose the column to split by.
5. Click **Segregate (FAST)**.
6. Open the new file saved as `yourfile_segregated.xlsx`.

---

## Key features

* Fast segregation using Pandas and XlsxWriter
* Select any sheet in the workbook
* Select any column as the grouping key
* Creates separate sheets for each unique value
* Keeps the original column order
* Shows progress while processing
* Handles long or duplicate sheet names safely

---

## Screenshots


![Main screen](https://github.com/metsploit/Excel-SheetSlicer/blob/main/Screenshot%202026-06-21%20231849.png)

![Sheet and column selection](https://github.com/metsploit/Excel-SheetSlicer/blob/main/Screenshot%202026-06-21%20231943.png)

![Output result](https://github.com/metsploit/Excel-SheetSlicer/blob/main/Screenshot%202026-06-21%20232110.png)

---

## Install

```bash
git clone https://github.com/yourusername/excel-sheet-segregator-fast.git
cd excel-sheet-segregator-fast
pip install pandas openpyxl XlsxWriter
python "Fast good for large data but no format.py"
```

---

## Best for

* Operations reports
* Region-wise reports
* Vendor-wise files
* Category-wise splitting
* Large trackers and master sheets
# Excel-SheetSlicer
Split one Excel sheet into multiple sheets by any column value.
