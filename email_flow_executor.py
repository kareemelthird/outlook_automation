import win32com.client as win32
import pythoncom  # Import pythoncom
import json
import os
from table_extraction import extract_table_content
from datetime import datetime

def process_placeholders(text):
    now = datetime.now()
    replacements = {
        '{today}': now.strftime('%Y-%m-%d'),
        '{now}': now.strftime('%Y-%m-%d %H:%M:%S')
    }
    for placeholder, value in replacements.items():
        text = text.replace(placeholder, value)
    return text


def send_email_via_outlook(flow):
    # Initialize COM for this thread
    pythoncom.CoInitialize()

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0 indicates a mail item
    mail.To = flow["to_email"]
    if flow["cc_email"]:
        mail.CC = flow["cc_email"]

    processed_subject = process_placeholders(flow["subject"])
    processed_body = process_placeholders(flow["body"])

    # Handle table content replacement
    for table_identifier in flow["tables_to_extract"]:
        parts = table_identifier.split(" - ")
        if len(parts) == 3:
            filename, sheetname, tablename = parts
            placeholder = f"[TABLE: {filename} - {tablename}]"
            if placeholder in processed_body:
                html_table_content = extract_table_content(flow["attachments"], table_identifier)
                processed_body = processed_body.replace(placeholder, html_table_content)

    mail.Subject = processed_subject
    mail.HTMLBody = processed_body

    # Add attachments
    for attachment in flow["attachments"].split(", "):
        if attachment.strip():
            mail.Attachments.Add(attachment.strip())

    mail.Send()    # Initialize COM for this thread
    pythoncom.CoInitialize()

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # 0 indicates a mail item
        mail.To = flow["to_email"]
        if flow["cc_email"]:
            mail.CC = flow["cc_email"]
        mail.Subject = process_placeholders(flow["subject"])

        html_body = process_placeholders(flow["body"])

        for table_identifier in flow["tables_to_extract"]:
            parts = table_identifier.split(" - ")
            if len(parts) == 3:
                filename, sheetname, tablename = parts
                placeholder = f"[TABLE: {filename} - {tablename}]"
                if placeholder in html_body:
                    html_table_content = extract_table_content(flow["attachments"], table_identifier)
                    html_body = html_body.replace(placeholder, html_table_content)

        mail.HTMLBody = html_body

        # Add attachments
        for attachment in flow["attachments"].split(", "):
            if attachment.strip():  # Avoid adding empty strings as attachments
                mail.Attachments.Add(attachment.strip())

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
    flow_path = 'email_flows/dd.json'  # Replace with your actual flow file
    if os.path.exists(flow_path):
        with open(flow_path, 'r') as f:
            flow = json.load(f)
        send_email_via_outlook(flow)
