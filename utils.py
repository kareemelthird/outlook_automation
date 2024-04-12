# utils.py
import os
import tkinter as tk
from tkinter import messagebox

def update_listbox(listbox, flow_folder):
    listbox.delete(0, tk.END)
    flow_names = [os.path.splitext(file)[0] for file in os.listdir(flow_folder) if file.endswith('.json')]
    for name in flow_names:
        listbox.insert(tk.END, name)
