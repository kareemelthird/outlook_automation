import logging
import os
import time
import json
from datetime import datetime
from pystray import MenuItem as item, Icon
from PIL import Image
import pythoncom
import email_flow_executor
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from plyer import notification
import tkinter as tk
from home import HomeApp
from ui_triggered_flow import load_triggered_flow_config

# Setup logging
logging.basicConfig(filename='email_automation.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

def open_home_screen():
    global root  # To keep a reference to your Tkinter root window
    if 'root' in globals():
        root.deiconify()  # Brings the window to the front if minimized or hidden
    else:
        root = tk.Tk()  # Create a new Tk root if it does not exist
        root.protocol("WM_DELETE_WINDOW", hide_window)  # Intercept close button
        app = HomeApp(root)  # Assuming HomeApp is your main application class
        root.mainloop()

def hide_window():
    global root
    root.withdraw()  # Hide the window instead of closing

class FolderWatcher:
    def __init__(self, folder_path, email_details, flow_name):
        self.observer = Observer()
        self.folder_path = folder_path
        self.email_details = email_details
        self.flow_name = flow_name
        self.ensure_directory_exists()

    def ensure_directory_exists(self):
        if not os.path.exists(self.folder_path):
            logging.info(f"Folder '{self.folder_path}' not found. Creating...")
            os.makedirs(self.folder_path, exist_ok=True)

    def run(self):
        event_handler = Handler(self.email_details, self.flow_name)
        self.observer.schedule(event_handler, self.folder_path, recursive=True)
        self.observer.start()

class Handler(FileSystemEventHandler):
    def __init__(self, email_details, flow_name):
        self.email_details = email_details
        self.flow_name = flow_name

    def on_created(self, event):
        if not event.is_directory:
            # Log details about the flow instead of just the file detection
            log_details = (
                f"Triggered Flow Detected: {self.flow_name}, "
                f"Type: Triggered, "
                f"To: {self.email_details['to_email']}, "
                f"CC: {self.email_details.get('cc_email', 'N/A')}, "
                f"Attachments: {self.email_details.get('attachments', 'None')}"
            )
            logging.info(log_details)
            pythoncom.CoInitialize()
            try:
                email_flow_executor.send_email_via_outlook(
                    flow=self.email_details,
                    additional_attachment=event.src_path
                )
                notify_user(f"Triggered email sent for '{self.flow_name}' flow.")
            except Exception as e:
                logging.error(f"Type: Scheduled Error sending triggered email for '{self.flow_name}': {e}")
            finally:
                pythoncom.CoUninitialize()


def log_flow_details(flow_name, flow_data):
    logging.info(f"Flow Details - Name: {flow_name}, Type: {flow_data.get('type', 'N/A')}, "
                 f"To: {flow_data.get('to_email', 'N/A')}, CC: {flow_data.get('cc_email', 'N/A')}, "
                 f"Subject: {flow_data.get('subject', 'N/A')}, Attachments: {flow_data.get('attachments', 'N/A')}")


def notify_user(message, title='K3 Email Automation Tool'):
    notification.notify(
        title=title,
        message=message,
        app_name='K3 Email Automation Tool',
        app_icon='k3.ico'  # Ensure the icon path is correct
    )
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
            try:
                with open(os.path.join(flow_folder, flow_file), 'r') as f:
                    flow = json.load(f)
                if is_time_to_send(flow, current_time, current_day_index):
                    log_details = (
                        f"Scheduled Flow Detected: {os.path.splitext(flow_file)[0]}, "
                        f"Type: Scheduled, "
                        f"To: {flow['to_email']}, "
                        f"CC: {flow.get('cc_email', 'N/A')}, "
                        f"Attachments: {flow.get('attachments', 'None')}, "
                        f"Tables: {flow.get('tables_to_extract', 'None')}, "
                        f"Scheduled Time: {flow['scheduled_time']}"
                    )
                    logging.info(log_details)
                    retry_send_email(flow, flow_file, 3)  # Retry sending email up to 3 times
            except json.JSONDecodeError as e:
                logging.error(f"Type: Scheduled Failed to parse JSON from {flow_file}: {e}")
            except Exception as e:
                logging.error(f"Unhandled exception for {flow_file}: {e}")


def retry_send_email(flow, flow_file, retries):
    flow_name = os.path.splitext(flow_file)[0]  # Extract flow name without extension
    while retries > 0:
        try:
            email_flow_executor.execute_flow(flow_file)
            logging.info(f"Schedule Email sent successfully for '{flow_name}', retries left: {retries}")
            notify_user(f"Schedule Email successfully sent for '{flow_name}'.", title="Schedule Email Sent Notification")
            break
        except Exception as e:
            logging.error(f"Type: Scheduled Failed to schedule send email for {flow_name}, error: {e}, retries left: {retries - 1}")
            retries -= 1
            time.sleep(10)  # Wait for 10 seconds before retrying

    if retries == 0:
        error_message = f"Type: Scheduled Exhausted retries for sending schedule email for {flow_name}"
        logging.error(error_message)
        notify_user(error_message, title="Schedule Email Sending Failed")


def on_clicked(icon, item):
    icon.stop()
    os._exit(1)
if 'root' in globals():
    root.withdraw()  # Hide the main window instead of stopping the application
    # Do not call icon.stop() unless you want to actually exit the app
def setup(icon):
    pythoncom.CoInitialize()  # Initialize COM
    icon.visible = True

    # Add a menu item for opening the home screen
    menu_items = (
        item('K3 Outlook Automation Home', open_home_screen),  # Calls open_home_screen when selected
        item('Exit', on_clicked)
    )
    icon.menu = menu_items

    start_folder_watcher()  # Start watching folders

    try:
        while True:
            check_and_send_emails(icon)
            time.sleep(60)  # Check every minute
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    finally:
        pythoncom.CoUninitialize()

def start_folder_watcher():
    config = load_triggered_flow_config()
    for filename, flow_config in config.items():
        folder = flow_config.get('folder', '')
        email_details = flow_config.get('email_details', {})
        flow_name = os.path.splitext(filename)[0]
        if folder:
            watcher = FolderWatcher(folder, email_details, flow_name)
            watcher.run()
        else:
            print(f"Folder path not provided for flow '{flow_name}'. Skipping setup.")
            
image = Image.open("k3.ico")  # Load the icon image
menu = (item('Exit', on_clicked),)
icon = Icon("K3 Email Automation Tool", image, "K3 Email Automation Tool", menu)
icon.run(setup)
