# merge-excel-sheets-into-one

🧾 Excel Sheet Merger (by Sheet Index)

A simple Python script to merge multiple Excel sheets (from a specific index range) into one combined sheet — without needing to specify each sheet name.
Supports command-line arguments and can handle large workbooks efficiently.

🚀 Features

🧩 Merge sheets by index range (e.g., from sheet #12 to #253)

🗂️ Keeps header from the first sheet only

⚙️ Command-line support with -i, -o, --start, --end options

💾 Saves the result as a single Excel file

🧱 Works with large .xlsx files

🧰 Requirements

Python 3.8+

Package: openpyxl

Install the required package:

pip install openpyxl

📄 Usage
Basic Example
python merge_sheets_by_index.py -i "C:\Users\You\Documents\Gazette11th_2025.xlsx" -o "C:\Users\You\Documents\Combined_Tables.xlsx"


This merges all sheets from index 12 to index 253 (default range).

Custom Range
python merge_sheets_by_index.py -i "file.xlsx" -o "combined.xlsx" --start 10 --end 50


Merges sheets from #10 through #50 only.

Show Help
python merge_sheets_by_index.py -h


Output:

usage: merge_sheets_by_index.py [-h] -i INPUT -o OUTPUT [--start START] [--end END]

Merge Excel sheets by index range into a single sheet.

options:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Path to the input Excel file
  -o OUTPUT, --output OUTPUT
                        Path to save the combined Excel file
  --start START         Starting sheet index (default: 12)
  --end END             Ending sheet index (default: 253)

🧮 Example Workflow

You have an Excel file with hundreds of sheets (Gazette11th_2025.xlsx)

You only want to merge sheets 12 to 253 into a single sheet

Run:

python merge_sheets_by_index.py -i Gazette11th_2025.xlsx -o Combined.xlsx


Result: Combined.xlsx containing all selected sheets combined into one.

🧱 Script File Structure
project-folder/
├── merge_sheets_by_index.py
└── README.md

🧑‍💻 Example Output

After running, you’ll see something like:

🔄 Loading workbook: Gazette11th_2025.xlsx
📘 Creating new workbook...
📄 Processing sheet #12: Table 12
📄 Processing sheet #13: Table 13
...
💾 Saving combined workbook...
✅ Done! Combined file saved at: Combined_Tables.xlsx
