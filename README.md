PDF Table Extraction and Export to Excel
Overview
This tool extracts tables from a PDF file and saves them into an Excel file, with each table stored in a separate sheet. It uses pdfplumber for table detection, pandas for data handling, and openpyxl for writing to Excel.

Dependencies & Installation
Ensure you have Python 3.x installed, then install the required libraries:

bash
Copy
Edit
pip install pdfplumber pandas openpyxl
How to Run
Execute the script by running:
bash
Copy
Edit
python script.py
Select the PDF file when prompted.
Choose an output location for the Excel file.
The script will process the PDF and save extracted tables into an Excel file.
Features
✔ Detects tables on all pages of a PDF
✔ Saves each table in a separate Excel sheet
✔ Works with PDFs with or without table borders

Limitations
May not handle complex or irregular tables accurately.
Performance depends on the PDF's formatting and text structure.
Future Enhancements
Improved handling for non-bordered tables.
Support for multi-line cell content.
Preview extracted tables before saving.
