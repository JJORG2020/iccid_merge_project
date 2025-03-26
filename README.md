ICCID Merge & Clean Tool
=========================

This Python script merges data from multiple Excel export files, deduplicates entries based on ICCID, and compares them to a master database to remove matching rows.

Features:
---------
- Merges 18 `Export_#.xlsx` files
- Deduplicates by `ICCID`
- Compares to `Main_Data.xlsx`
- Outputs:
  - `Merged_Data.xlsx`: Combined ICCIDs
  - `Filtered_Main_Data.xlsx`: Main DB excluding matched ICCIDs
  - `Main_data_with_highlight.xlsx`: Main DB with matched rows highlighted in red

File Structure:
---------------
~/iccid_merge_project/
├── Export_1.xlsx → Export_18.xlsx
├── Main_Data.xlsx
├── Merged_Data.xlsx
├── Filtered_Main_Data.xlsx
├── Main_data_with_highlight.xlsx
├── iccid_merge.py
└── README.txt

How to Use:
-----------
1. Ensure all `Export_#.xlsx` files and `Main_Data.xlsx` are placed in `~/iccid_merge_project/`
2. Activate your virtual environment:
   source venv/bin/activate
3. Run the script:
   python3 iccid_merge.py

Requirements:
-------------
- Python 3.8+
- pandas
- openpyxl
- xlsxwriter

Install dependencies with:
pip install pandas openpyxl xlsxwriter

License:
--------
Private internal use only. Not licensed for redistribution.

Author:
-------
JJORG2020
