#ui_file_browsing.py
import os
from tkinter import filedialog, messagebox
import openpyxl

def browse_files(attachments_var, table_list_widget):
    filenames = filedialog.askopenfilenames(filetypes=[("All files", "*.*")])  # Allow all file types
    if filenames:
        current_files = attachments_var.get().split(", ") if attachments_var.get() else []
        updated_files = list(set(current_files + list(filenames)))  # Combine and remove duplicates

        attachments_var.set(", ".join(updated_files))
        update_table_list(table_list_widget, filenames)

def update_table_list(table_list_widget, filenames):
    for file in filenames:
        if file.endswith(('.xlsx', '.xls')):
            try:
                workbook = openpyxl.load_workbook(file, data_only=True)
                for worksheet in workbook.worksheets:
                    for table in worksheet.tables.values():
                        table_list_widget.insert('end', f"{os.path.basename(file)} - {worksheet.title} - {table.name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file tables: {e}")

