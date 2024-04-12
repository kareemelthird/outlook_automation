#table_extraction.py
import pandas as pd
import openpyxl
import os

def extract_table_content(attachments, table_identifier):
    html_content = ""
    attachment_paths = attachments.split(", ")
    for attachment in attachment_paths:
        if not attachment.endswith(('.xlsx', '.xls', '.xlsm', '.xltm')):
            continue  # Skip non-Excel files

        parts = table_identifier.split(" - ")
        if len(parts) < 3:
            print(f"Invalid table identifier format: {table_identifier}")
            continue

        tablename = parts[-1]
        sheetname = parts[-2]
        filename = " - ".join(parts[:-2])

        # Check if the file basename matches the filename in the identifier
        if os.path.basename(attachment) != filename:
            continue

        workbook = openpyxl.load_workbook(attachment, data_only=True)
        worksheet = workbook[sheetname]
        table = worksheet.tables[tablename]

        start_cell, end_cell = table.ref.split(":")
        start_col = ''.join(filter(str.isalpha, start_cell))
        end_col = ''.join(filter(str.isalpha, end_cell))
        start_row = ''.join(filter(str.isdigit, start_cell))

        data = pd.read_excel(attachment, sheet_name=sheetname, header=int(start_row)-1, engine='openpyxl', usecols=f"{start_col}:{end_col}")
        
        # Convert columns to strings to handle NaN replacements
        data = data.astype(str)
        data.replace('nan', '', inplace=True)

        html_content += data.to_html(index=False)

    return html_content
