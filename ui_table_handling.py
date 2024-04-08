from datetime import datetime
import tkinter as tk
from tkinter import messagebox

def insert_placeholder(placeholder, widget):
    if placeholder == 'today':
        placeholder_text = '{today}'
    elif placeholder == 'now':
        placeholder_text = '{now}'
    else:
        placeholder_text = ''  # You can extend this for other placeholders

    widget.insert(tk.INSERT, placeholder_text)


def insert_table_into_body(selected_items, body_text_widget):
    for selected_item in selected_items:
        try:
            file_name, worksheet_title, table_name = selected_item.split(" - ")
            placeholder = f"[TABLE: {file_name} - {table_name}]"
            body_text_widget.insert(tk.END, placeholder + '\n')
        except ValueError as ve:
            messagebox.showerror("Error", "Invalid table selection format")
        except Exception as e:
            messagebox.showerror("Error", f"Error inserting table: {e}")


