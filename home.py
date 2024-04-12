import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkinter import scrolledtext
from Missed_Runs import analyze_missed_executions
from main_ui import EmailAutomationApp  # Import Scheduled Email Flow Editor
from ui_triggered_flow import setup_triggered_flow_ui  # Import Triggered Email Flow Editor
from PIL import Image, ImageTk
import multiprocessing
import sys


def run_background_process():
    import subprocess
    # Assuming background.py is in the same directory
    subprocess.run([sys.executable, 'background.py'])


def load_and_filter_log_entries(log_file_path, flow_type):
    try:
        with open(log_file_path, 'r') as file:
            all_entries = file.readlines()
        # Filter entries by flow type
        filtered_entries = [entry for entry in all_entries if f"Type: {flow_type}" in entry]
        return filtered_entries
    except IOError as e:
        messagebox.showerror("Read Error", f"Failed to read log file: {e}")
        return []

class HomeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("K3 Automation Tool")
        self.root.geometry("880x660")
        self.background_process = multiprocessing.Process(target=run_background_process)
        self.background_process.start()

        # Main Frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        # Footer
        self.create_footer()

        # Title Label
        title_label = ttk.Label(self.main_frame, text="K3 Outlook Automation Tool", 
                                font=("Helvetica", 20, "bold"), foreground="#4CAF50")
        title_label.pack(pady=(0, 20))

        # Create Flow Management Sections
        self.create_flow_management_section('Scheduled', 'email_flows', 
                                            self.new_scheduled_flow, 
                                            self.edit_scheduled_flow, 
                                            self.delete_flow)
        self.create_flow_management_section('Triggered', 'trigger_flows', 
                                            self.new_triggered_flow, 
                                            self.edit_triggered_flow, 
                                            self.delete_flow)
        # Missed Runs Section
        self.create_missed_runs_section()

    def create_missed_runs_section(self):
        missed_runs_frame = ttk.LabelFrame(self.main_frame, text="Missed Runs")
        missed_runs_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.missed_runs_list = tk.Listbox(missed_runs_frame, height=5)
        self.missed_runs_list.pack(fill='both', expand=True, padx=10, pady=10)

        refresh_button = ttk.Button(missed_runs_frame, text="Refresh Missed Runs",
                                    command=self.refresh_missed_runs)
        refresh_button.pack(side='bottom', pady=5)

        self.refresh_missed_runs()

    def refresh_missed_runs(self):
        self.missed_runs_list.delete(0, tk.END)  # Clear existing entries
        missed_runs = analyze_missed_executions()
        for flow_name, dates in missed_runs.items():
            for date in dates:
                self.missed_runs_list.insert(tk.END, f"{flow_name}: Missed on {date}")

    


    def create_footer(self):
        footer_frame = ttk.Frame(self.main_frame)
        footer_frame.pack(side='bottom', fill='x', pady=5)

        # Load and display the icon
        self.icon_image = ImageTk.PhotoImage(Image.open('k3.ico'))
        icon_label = ttk.Label(footer_frame, image=self.icon_image)
        icon_label.pack(side='left', padx=10)

        # Your professional details
        details_label = ttk.Label(footer_frame, text="Developed By: K3\nLinkedIn: Kareem Hassan",
                                  font=("Helvetica", 10), foreground="blue")
        details_label.pack(side='left')

        # Make the URL label clickable if desired
        details_label.bind("<Button-1>", lambda e: self.open_link("https://www.linkedin.com/signup/public-profile-join?vieweeVanityName=kareemelthird&trk=public_profile-settings_top-card-primary-button-join-to-view-profile"))

    def open_link(self, url):
        import webbrowser
        webbrowser.open_new(url)


    def create_flow_management_section(self, flow_type, folder, new_function, edit_function, delete_function):
        # Main container for flow management and log
        main_container = ttk.Frame(self.main_frame)
        main_container.pack(side='top', fill='both', expand=True, padx=10, pady=10)

        # Management frame
        management_frame = ttk.LabelFrame(main_container, text=f"{flow_type} Flow Management")
        management_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)



        lb_flows = tk.Listbox(management_frame, height=5)
        lb_flows.pack(pady=5, padx=5, fill='both', expand=True)
        lb_flows.bind('<<ListboxSelect>>', lambda event: self.refresh_log(event, log_text, folder))  # Bind selection event


        btn_frame = ttk.Frame(management_frame)
        btn_frame.pack(fill='x', pady=5)

        btn_new = ttk.Button(btn_frame, text="New", command=lambda: new_function(folder, lb_flows))
        btn_new.pack(side='left', fill='x', expand=True)

        btn_edit = ttk.Button(btn_frame, text="Edit", command=lambda: edit_function(lb_flows.get(tk.ACTIVE), folder, lb_flows))
        btn_edit.pack(side='left', fill='x', expand=True)

        btn_delete = ttk.Button(btn_frame, text="Delete", command=lambda: delete_function(lb_flows.get(tk.ACTIVE), folder, lb_flows))
        btn_delete.pack(side='left', fill='x', expand=True)

        self.update_listbox(lb_flows, folder)


        # Log frame
        log_frame = ttk.LabelFrame(main_container, text=f"{flow_type} Log")
        log_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)

        # Create a ScrolledText widget for log entries
        log_text = scrolledtext.ScrolledText(log_frame, height=10, width=50)
        log_text.pack(fill='both', expand=True)

        # Button frame for log actions
        log_button_frame = ttk.Frame(log_frame)
        log_button_frame.pack(fill='x', pady=5)



        # Initially populate the log panel with entries
        self.refresh_log(None, log_text, folder)


        # Load log entries
        log_file_path = 'email_automation.log'  # Update with the correct log file path
        log_entries = load_and_filter_log_entries(log_file_path, flow_type)
        for entry in log_entries:
            log_text.insert(tk.END, entry)
        log_text.configure(state='disabled')  # Make the log panel read-only

    def refresh_log(self, event, log_text_widget, folder):
        # Determine flow type based on the folder
        flow_type = "Triggered" if folder == 'trigger_flows' else "Scheduled"

        # When called without an event, it means refresh all for flow_type
        selected_workflow = None
        if event:  # Called from a listbox selection event
            selection = event.widget.curselection()
            if selection:
                selected_workflow = event.widget.get(selection[0])

        log_text_widget.configure(state='normal')
        log_text_widget.delete('1.0', tk.END)
        log_entries = load_and_filter_log_entries('email_automation.log', flow_type)
        for entry in log_entries:
            # Insert log entry if no specific workflow is selected, or if it matches the selected workflow
            if selected_workflow is None or selected_workflow in entry:
                log_text_widget.insert(tk.END, entry)
        log_text_widget.configure(state='disabled')


    
    def new_scheduled_flow(self, folder, listbox):
        # Open the EmailAutomationApp for creating a new flow
        EmailAutomationApp(tk.Toplevel(), None, None, listbox, folder)

    def new_triggered_flow(self, folder, listbox):
        setup_triggered_flow_ui(None, listbox)

    def edit_scheduled_flow(self, flow_name, folder, listbox):
        if flow_name:
            settings_path = os.path.join(folder, f"{flow_name}.json")
            with open(settings_path, 'r') as file:
                settings = json.load(file)
            EmailAutomationApp(tk.Toplevel(), settings, flow_name, listbox, folder)

    def edit_triggered_flow(self, flow_name, folder, listbox):
        if flow_name:
            setup_triggered_flow_ui(flow_name, listbox)

    def delete_flow(self, flow_name, folder, listbox):
        if flow_name and messagebox.askyesno("Delete Flow", f"Are you sure you want to delete '{flow_name}'?"):
            os.remove(os.path.join(folder, f"{flow_name}.json"))
            self.update_listbox(listbox, folder)

    def update_listbox(self, listbox, folder):
        listbox.delete(0, tk.END)
        for file in os.listdir(folder):
            if file.endswith(".json"):
                listbox.insert(tk.END, file[:-5])

def launch_home_ui():
    root = tk.Tk()
    HomeApp(root)
    root.mainloop()

if __name__ == "__main__":
    launch_home_ui()
    