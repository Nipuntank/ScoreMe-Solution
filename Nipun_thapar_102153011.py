import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog


def detect_tables(page):
    """
    Detect tables in a PDF page by analyzing text and layout.
    Handles tables with borders, no borders, and irregular shapes.
    """
    tables = []
    words = page.extract_words(
        x_tolerance=2, y_tolerance=2
    )  # Extract words with tolerances
    lines = page.lines  # Extract lines from the page

    rows = {}  # Dictionary to store words grouped by y-coordinates
    for word in words:
        y = word["top"]  # Get y-coordinate of the word
        if y not in rows:
            rows[y] = []  # Initialize list if not exists
        rows[y].append(word)  # Append word to its respective row

    sorted_rows = sorted(rows.items(), key=lambda x: x[0])  # Sort rows by y-coordinate

    table_data = []  # Initialize list to store table data
    for y, words_in_row in sorted_rows:
        row = []
        for word in sorted(
            words_in_row, key=lambda x: x["x0"]
        ):  # Sort words by x-coordinates
            row.append(word["text"])  # Append word text to row
        table_data.append(row)

    max_cols = max(len(row) for row in table_data)  # Determine max columns in a row
    for row in table_data:
        while len(row) < max_cols:
            row.append("")  # Fill empty cells for consistency

    tables.append(table_data)
    return tables  # Return detected tables


def save_to_excel(tables, output_file):
    """
    Save extracted tables to an Excel file.
    Each table is saved as a separate sheet.
    """
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for i, table in enumerate(tables):
            df = pd.DataFrame(table)  # Convert table to DataFrame
            df.to_excel(
                writer, sheet_name=f"Table_{i+1}", index=False
            )  # Save each table as a separate sheet


def extract_tables_from_pdf(pdf_path, output_excel_path):
    """
    Extract tables from a PDF and save them to an Excel file.
    """
    tables = []  # List to store extracted tables
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Processing page {page_num + 1}...")  # Log progress
            page_tables = detect_tables(page)  # Detect tables in the page
            for table in page_tables:
                tables.append(table)  # Append extracted tables

    save_to_excel(tables, output_excel_path)  # Save all tables to Excel
    print(f"Tables extracted and saved to {output_excel_path}")


def select_file():
    """
    Open a file picker dialog to select the input PDF file.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select a PDF file", filetypes=[("PDF Files", "*.pdf")]
    )
    return file_path  # Return selected file path


def main():
    """
    Main function to run the script with a file picker for PDF input and Excel output.
    """
    pdf_path = select_file()  # Open file dialog for PDF selection
    if not pdf_path:
        print("No file selected. Exiting...")
        return

    output_excel_path = filedialog.asksaveasfilename(
        title="Save Excel file as",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
    )
    if not output_excel_path:
        print("No output file selected. Exiting...")
        return

    extract_tables_from_pdf(pdf_path, output_excel_path)  # Extract and save tables


if __name__ == "__main__":
    main()  # Execute the script
