import win32com.client as win32
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
