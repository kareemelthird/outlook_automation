import win32com.client as win32
import pythoncom
from table_extraction import extract_table_content
from datetime import datetime
import json
import os

def process_placeholders(text):
    now = datetime.now()
    replacements = {
        '{today}': now.strftime('%Y-%m-%d'),
        '{now}': now.strftime('%Y-%m-%d %H:%M:%S')
    }
    for placeholder, value in replacements.items():
        text = text.replace(placeholder, value)
    return text

def send_email_via_outlook(flow, additional_attachment=None):
    # Initialize COM for this thread
    pythoncom.CoInitialize()

    try:
        # Create Outlook Mail Item
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = flow["to_email"]
        if flow.get("cc_email"):
            mail.CC = flow["cc_email"]
        mail.Subject = process_placeholders(flow["subject"])
        mail.HTMLBody = process_placeholders(flow["body"])

        # Process tables_to_extract if available
        if 'tables_to_extract' in flow:
            for table_identifier in flow["tables_to_extract"]:
                parts = table_identifier.split(" - ")
                if len(parts) == 3:
                    filename, sheetname, tablename = parts
                    placeholder = f"[TABLE: {filename} - {tablename}]"
                    if placeholder in mail.HTMLBody:
                        html_table_content = extract_table_content(flow["attachments"], table_identifier)
                        mail.HTMLBody = mail.HTMLBody.replace(placeholder, html_table_content)

        # Add regular attachments from flow
        if "attachments" in flow:
            for attachment in flow["attachments"].split(", "):
                if attachment.strip():
                    mail.Attachments.Add(attachment.strip())

        # Add additional_attachment for triggered emails
        if additional_attachment:
            mail.Attachments.Add(additional_attachment)

        # Send the email
        mail.Send()

    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()

def execute_flow(flow_file_name):
    flow_folder = 'email_flows'
    flow_path = os.path.join(flow_folder, flow_file_name)
    if os.path.exists(flow_path):
        with open(flow_path, 'r') as f:
            flow = json.load(f)
        send_email_via_outlook(flow)

if __name__ == "__main__":
    flow_path = 'email_flows/1.json'  # Replace with your actual flow file
    if os.path.exists(flow_path):
        with open(flow_path, 'r') as f:
            flow = json.load(f)
        send_email_via_outlook(flow)
