#main_ui.py
import os
import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import simpledialog
from tkinter import messagebox
from ui_table_handling import insert_table_into_body
from ui_submit import submit_action

from ui_file_browsing import browse_files



def browse_attachments(self):
    browse_files(self.attachments, self.table_list)
    self.attachments_entry.insert(0, self.attachments.get())




class EmailAutomationApp:
    def __init__(self, root, settings=None, flow_name=None, listbox_widget=None, folder='email_flows'):
        self.root = root
        self.root.title("K3 Automation Tool")
        self.root.geometry("680x660")

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(padx=20, pady=20, fill='both', expand=True)

        self.create_header()
        self.create_email_details()
        self.create_schedule()
        self.create_submit_button()
        

        # Set the current_flow_name attribute
        self.current_flow_name = flow_name
        self.listbox_widget = listbox_widget
        self.folder = folder

        if settings:
            self.load_settings(settings, flow_name)



    def submit_action(self):
        # Gathering email details from the form
        selected_days = [day.get() for day in self.days]
        email_details = {
            "to_email": self.to_email_entry.get(),
            "cc_email": self.cc_email_entry.get(),
            "subject": self.subject_entry.get(),
            "body": self.body_text.get("1.0", tk.END),
            "attachments": self.attachments.get(),
            "selected_days": selected_days,
            "scheduled_time": self.time_entry.get(),
            "tables_to_extract": self.selected_table_identifiers
        }

        # Use the current flow name for saving; no new flow name prompt
        submit_action(
            email_details["to_email"],
            email_details["cc_email"],
            email_details["subject"],
            email_details["body"],
            email_details["attachments"],
            email_details["selected_days"],
            email_details["scheduled_time"],
            email_details["tables_to_extract"],
            flow_name=self.current_flow_name  # Uses the existing flow name
        )
        # Update listbox if the widget is provided
        if self.listbox_widget:
            self.update_listbox_after_changes()
            
        # Close the EmailAutomationApp window after successful submission
        self.root.destroy()



    def ask_for_flow_name(self):
        return simpledialog.askstring("Flow Name", "Enter a name for this email flow:")


    def update_listbox_after_changes(self):
        self.listbox_widget.delete(0, tk.END)
        for file in os.listdir(self.folder):
            if file.endswith(".json"):
                self.listbox_widget.insert(tk.END, file[:-5])

    # Add a delete flow function if you have a delete option within the UI
    def delete_flow(self, flow_name):
        if flow_name and messagebox.askyesno("Delete Flow", f"Are you sure you want to delete '{flow_name}'?"):
            os.remove(os.path.join(self.folder, f"{flow_name}.json"))
            self.update_listbox_after_changes()



    def create_header(self):
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, columnspan=2, sticky='ew')
        header_label = ttk.Label(header_frame, text="K3 Outlook Automation - Scheduled Flow", font=("Helvetica", 20, "bold"), foreground="#4CAF50")
        header_label.pack(pady=(0, 20))

    def create_email_details(self):
        email_frame = ttk.LabelFrame(self.main_frame, text="Email Details")
        email_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # To Email
        ttk.Label(email_frame, text="To Email:", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.to_email_entry = ttk.Entry(email_frame, width=80)
        self.to_email_entry.grid(row=0, column=1, padx=5, pady=5)

        # CC Email
        ttk.Label(email_frame, text="CC Email:", font=("Helvetica", 10)).grid(row=1, column=0, padx=5, pady=5)
        self.cc_email_entry = ttk.Entry(email_frame, width=80)
        self.cc_email_entry.grid(row=1, column=1, padx=5, pady=5)

        # Subject
        ttk.Label(email_frame, text="Subject:", font=("Helvetica", 10)).grid(row=2, column=0, padx=5, pady=5)
        self.subject_entry = ttk.Entry(email_frame, width=80)
        self.subject_entry.grid(row=2, column=1, padx=5, pady=5)

        # Body
        ttk.Label(email_frame, text="Body:", font=("Helvetica", 10)).grid(row=3, column=0)
        self.body_text = scrolledtext.ScrolledText(email_frame, height=4, width=60)
        self.body_text.grid(row=3, column=1, padx=5, pady=5)

        # Attachments and Tables Frame
        attachments_frame = ttk.LabelFrame(email_frame, text="Attachments", padding=10)
        attachments_frame.grid(row=4, columnspan=2, pady=10, sticky="ew")

        # Attachments
        ttk.Label(attachments_frame, text="Attachments:", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.attachments = tk.StringVar()
        self.attachments_entry = ttk.Entry(attachments_frame, textvariable=self.attachments, width=60)
        self.attachments_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(attachments_frame, text="Browse", command=self.browse_attachments).grid(row=0, column=2, padx=5, pady=5)

        # Table List
        ttk.Label(attachments_frame, text="Tables:", font=("Helvetica", 10)).grid(row=1, column=0, padx=5, pady=5)
        self.table_list = tk.Listbox(attachments_frame, height=4, width=60, selectmode='multiple')
        self.table_list.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(attachments_frame, text="Insert Table", command=self.insert_table).grid(row=1, column=2, padx=5, pady=5)

    def create_schedule(self):
        schedule_frame = ttk.LabelFrame(self.main_frame, text="Schedule")
        schedule_frame.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        # Scheduled Time
        ttk.Label(schedule_frame, text="Scheduled Time (HH:MM):", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.time_entry = ttk.Entry(schedule_frame, width=70)
        self.time_entry.grid(row=0, column=1, padx=5, pady=5)

        # Weekdays
        days_frame = ttk.LabelFrame(schedule_frame, text="Weekdays", width=70)
        days_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="w")

        self.days = [tk.IntVar() for _ in range(7)]
        weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        for i, day in enumerate(weekdays):
            ttk.Checkbutton(days_frame, text=day, variable=self.days[i]).grid(row=0, column=i, padx=5, pady=5)

        self.selected_table_identifiers = []

    def create_submit_button(self):
        submit_button = ttk.Button(self.main_frame, text="Submit", command=self.submit_action)
        submit_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

    def browse_attachments(self):
        browse_files(self.attachments, self.table_list)

    def insert_table(self):
        selected_indices = self.table_list.curselection()
        print("Selected indices:", selected_indices)  # Debug print

        selected_items = [self.table_list.get(idx) for idx in selected_indices]
        print("Selected items:", selected_items)  # Debug print

        for item in selected_items:
            if item not in self.selected_table_identifiers:
                self.selected_table_identifiers.append(item)
                insert_table_into_body([item], self.body_text)  # Insert one item at a time
            else:
                messagebox.showinfo("Info", f"{item} is already added.")




    def load_settings(self, settings, flow_name=None):
        # Populating fields from settings dictionary
        self.to_email_entry.delete(0, tk.END)
        self.to_email_entry.insert(0, settings.get('to_email', ''))

        self.cc_email_entry.delete(0, tk.END)
        self.cc_email_entry.insert(0, settings.get('cc_email', ''))

        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, settings.get('subject', ''))


        self.body_text.delete('1.0', tk.END)
        self.body_text.insert(tk.INSERT, settings.get('body', ''))

        self.attachments.set(settings.get('attachments', ''))

        for idx, day in enumerate(settings.get('selected_days', [])):
            self.days[idx].set(day)

        self.time_entry.delete(0, tk.END)
        self.time_entry.insert(0, settings.get('scheduled_time', ''))

    # Load tables_to_extract if available
        self.table_list.delete(0, tk.END)  # Clear existing entries
        self.selected_table_identifiers = settings.get('tables_to_extract', [])
        for table_identifier in self.selected_table_identifiers:
            self.table_list.insert(tk.END, table_identifier)
        self.current_flow_name = flow_name  # Set current_flow_name for existing flow


# Main function to launch the UI
def launch_ui():
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()

if __name__ == "__main__":
    launch_ui()
