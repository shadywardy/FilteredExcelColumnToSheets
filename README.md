# Excel Processor

## Overview
This Python-based tool allows you to filter and save data from Excel files (.xls and .xlsx formats) by specifying a column to filter. It reads an input Excel file, filters data based on a specific column, and generates separate Excel files for each unique value in that column.

## Features
- **Load Excel Files**: Load Excel files and process them in the specified format (`.xls` or `.xlsx`).
- **Filter by Column**: Filter the data based on unique values from a column of your choice.
- **Save Filtered Files**: Save the filtered data to new Excel files, named based on the filter column values.
- **Supports `.xls` and `.xlsx` Formats**: Choose between `.xls` or `.xlsx` output files.



![image](https://github.com/user-attachments/assets/d82b35c2-2b9c-47c3-9de4-e9ec0983089a)




## Requirements
- Python 3.9.13 (if running the script version)
- Required Python libraries (if running the script):
  - `pandas`
  - `openpyxl`

## Installation
If using the Python script:
1. Install Python 3.9.13 if not already installed.
2. Install dependencies using pip:
   ```sh
   pip install pandas openpyxl
   ```
3. Run the script:
   ```sh
   python excel_processor.py
   ```

If using the executable (`ExcelProcessor.exe`):
1. Download `ExcelProcessor.exe`.
2. Double-click to run the application.

## Usage
1. Place the input Excel file in the same directory as the script/executable.
2. Run the script or executable.
3. The processed Excel file will be generated automatically.

## Notes
- Ensure that the input Excel file follows the expected format.
- If encountering issues, check dependencies and Python version.


