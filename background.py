import json
import os
import time
from datetime import datetime
from pystray import MenuItem as item
import pystray
from PIL import Image, ImageDraw
from plyer import notification
import email_flow_executor  # Your module for sending emails


def day_of_week_to_int(day):
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    return days.index(day)

def is_time_to_send(flow, current_time, current_day_index):
    scheduled_time = flow.get("scheduled_time", "")
    selected_days = flow.get("selected_days", [])

    is_day_to_send = selected_days[current_day_index] == 1 if selected_days else False
    is_time_to_send = current_time == scheduled_time

    return is_day_to_send and is_time_to_send

def check_and_send_emails(icon):
    flow_folder = 'email_flows'
    current_time = datetime.now().strftime("%H:%M")
    current_day_index = day_of_week_to_int(datetime.now().strftime("%a"))

    for flow_file in os.listdir(flow_folder):
        if flow_file.endswith('.json'):
            with open(os.path.join(flow_folder, flow_file), 'r') as f:
                flow = json.load(f)
            if is_time_to_send(flow, current_time, current_day_index):
                print(f"Sending email for flow: {flow_file}")
                email_flow_executor.execute_flow(flow_file)
                notify_user(f"Sent email for flow: {os.path.splitext(flow_file)[0]}")  # Remove ".json"
            else:
                print(f"Not time to send for flow: {flow_file}")
    icon.update_menu()

def notify_user(message):
    notification.notify(
        title='K3 Email Automation Tool',
        message=message,
        app_name='K3 Email Automation Tool',
        app_icon='k3.ico'  # Path to your icon file
    )

def on_clicked(icon, item):
    icon.stop()
    os._exit(1)

def setup(icon):
    icon.visible = True
    while True:
        check_and_send_emails(icon)
        time.sleep(60)  # Check every minute

image = Image.open("k3.ico")  # Load the icon image
menu = (item('Exit', on_clicked),)
icon = pystray.Icon("K3 Email Automation Tool", image, "K3 Email Automation Tool", menu)
icon.run(setup)
