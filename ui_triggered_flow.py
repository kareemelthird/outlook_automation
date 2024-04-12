import json
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, simpledialog
import os
from utils import update_listbox

def save_triggered_flow_config(root, filename, folder_path, to_email, cc_email, subject, body):
    config = {
        'folder': folder_path,
        'email_details': {
            'to_email': to_email,
            'cc_email': cc_email,
            'subject': subject,
            'body': body
        }
    }
    flow_folder = 'trigger_flows'
    os.makedirs(flow_folder, exist_ok=True)
    file_path = os.path.join(flow_folder, f'{filename}.json')

    with open(file_path, 'w') as file:
        json.dump(config, file, indent=4)

    root.destroy()  # Close the window
    messagebox.showinfo("Success", f"Triggered flow configuration saved as '{filename}.json'.")


def load_triggered_flow_config(flow_name=None):
    config_folder = 'trigger_flows'
    if flow_name:
        file_path = os.path.join(config_folder, f'{flow_name}.json')
        if os.path.exists(file_path):
            with open(file_path, 'r') as file:
                return json.load(file)
        return None
    else:
        config = {}
        if os.path.exists(config_folder) and os.path.isdir(config_folder):
            for filename in os.listdir(config_folder):
                if filename.endswith('.json'):
                    file_path = os.path.join(config_folder, filename)
                    with open(file_path, 'r') as file:
                        flow_config = json.load(file)
                        config[filename] = flow_config
        return config

def setup_triggered_flow_ui(edit_flow_name=None, listbox_widget=None):
    root = tk.Tk()
    root.title("Setup Triggered Email Flow")
    root.geometry("680x660")

    existing_config = load_triggered_flow_config(edit_flow_name) if edit_flow_name else None

    main_frame = ttk.Frame(root)
    main_frame.pack(padx=20, pady=20, fill='both', expand=True)

    # Header
    header_frame = ttk.Frame(main_frame)
    header_frame.grid(row=0, columnspan=2, sticky='ew')
    header_label = ttk.Label(header_frame, text="K3 Outlook Automation - Triggered Flow", font=("Helvetica", 20, "bold"), foreground="#4CAF50")
    header_label.pack(pady=(0, 20))

    # Folder selection
    ttk.Label(main_frame, text="Folder to Monitor:").grid(row=1, column=0, padx=10, pady=10)
    folder_path_entry = ttk.Entry(main_frame, width=60)
    folder_path_entry.grid(row=1, column=1, padx=10, pady=10)
    folder_path_entry.insert(0, existing_config['folder'] if existing_config else '')

    # Browse button
    def browse_folder():
        directory = filedialog.askdirectory()
        if directory:
            folder_path_entry.delete(0, tk.END)
            folder_path_entry.insert(0, directory)
    ttk.Button(main_frame, text="Browse", command=browse_folder).grid(row=1, column=2, padx=10, pady=10)

    # Email details
    def create_label_entry_pair(label_text, row, default_value=''):
        ttk.Label(main_frame, text=label_text).grid(row=row, column=0, padx=10, pady=10)
        entry = ttk.Entry(main_frame, width=60)
        entry.grid(row=row, column=1, padx=10, pady=10)
        entry.insert(0, default_value)
        return entry

    to_email_entry = create_label_entry_pair("To Email:", 2, existing_config['email_details']['to_email'] if existing_config else '')
    cc_email_entry = create_label_entry_pair("CC Email:", 3, existing_config['email_details']['cc_email'] if existing_config else '')
    subject_entry = create_label_entry_pair("Subject:", 4, existing_config['email_details']['subject'] if existing_config else '')

    # Body
    ttk.Label(main_frame, text="Body:").grid(row=5, column=0, padx=10, pady=10)
    body_text = scrolledtext.ScrolledText(main_frame, width=44, height=4)
    body_text.grid(row=5, column=1, padx=10, pady=10)
    body_text.insert(tk.INSERT, existing_config['email_details']['body'] if existing_config else '')

    # Save button
    def save_flow():
        filename = edit_flow_name if edit_flow_name else simpledialog.askstring("Flow Name", "Enter a name for the triggered flow:")
        if not filename:
            messagebox.showwarning("No Flow Name", "Flow name is required to save.")
            return

        save_triggered_flow_config(
            root,  # Pass the root window here
            filename,
            folder_path_entry.get(),
            to_email_entry.get(),
            cc_email_entry.get(),
            subject_entry.get(),
            body_text.get("1.0", tk.END)
        )
        if listbox_widget:  # Update listbox if provided
            update_listbox(listbox_widget, 'trigger_flows')

    save_button = ttk.Button(main_frame, text="Save Triggered Flow", command=save_flow)
    save_button.grid(row=7, columnspan=3, pady=20)

    # Pre-populate fields if editing an existing flow
    if existing_config:
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, existing_config['folder'])
        to_email_entry.delete(0, tk.END)
        to_email_entry.insert(0, existing_config['email_details']['to_email'])
        cc_email_entry.delete(0, tk.END)
        cc_email_entry.insert(0, existing_config['email_details']['cc_email'])
        subject_entry.delete(0, tk.END)
        subject_entry.insert(0, existing_config['email_details']['subject'])
        body_text.delete("1.0", tk.END)
        body_text.insert(tk.INSERT, existing_config['email_details']['body'])

    root.mainloop()

if __name__ == "__main__":
    setup_triggered_flow_ui()

