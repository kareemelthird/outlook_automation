import json
import os
import email_flow_executor
from tkinter import messagebox, simpledialog

def submit_action(to_email, cc_email, subject, body, attachments, selected_days, scheduled_time, selected_tables, flow_name=None):
    if not flow_name:
        flow_name = simpledialog.askstring("Flow Name", "Enter a name for this email flow:")
        if not flow_name:
            messagebox.showwarning("Not Saved", "Email flow not saved. Please provide a name.")
            return

    email_details = {
        "to_email": to_email,
        "cc_email": cc_email,
        "subject": subject,
        "body": body,
        "attachments": attachments,
        "selected_days": selected_days,
        "scheduled_time": scheduled_time,
        "tables_to_extract": selected_tables
    }

    save_flow(email_details, flow_name)

def save_flow(email_details, flow_name):
    flow_folder = 'email_flows'
    os.makedirs(flow_folder, exist_ok=True)
    flow_path = os.path.join(flow_folder, f"{flow_name}.json")
    with open(flow_path, 'w') as f:
        json.dump(email_details, f, indent=4)
    messagebox.showinfo("Saved", f"Email flow '{flow_name}' saved successfully.")
