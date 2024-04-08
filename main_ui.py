import tkinter as tk
from tkinter import ttk, scrolledtext
from ui_submit import submit_action
from ui_file_browsing import browse_files
from ui_table_handling import insert_table_into_body



class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("K3 Automation Tool")
        self.root.geometry("680x660")

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(padx=20, pady=20, fill='both', expand=True)

        self.create_header()
        self.create_email_details()
        self.create_schedule()
        self.create_submit_button()




    def create_header(self):
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, columnspan=2, sticky='ew')
        header_label = ttk.Label(header_frame, text="K3 Outlook Automation Tool", font=("Helvetica", 20, "bold"), foreground="#4CAF50")  # Customizing header label
        header_label.pack(pady=(0, 20))





    def create_email_details(self):
        email_frame = ttk.LabelFrame(self.main_frame, text="Email Details")
        email_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    

        # To Email
        ttk.Label(email_frame, text="To Email:", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.to_email_entry = ttk.Entry(email_frame, width=80)
        self.to_email_entry.grid(row=0, column=1, padx=5, pady=5)
        self.to_email_entry.insert(0, "k.hassan@modern-egypt.com")  # Default time or placeholder


        # CC Email
        ttk.Label(email_frame, text="CC Email:", font=("Helvetica", 10)).grid(row=1, column=0, padx=5, pady=5)
        self.cc_email_entry = ttk.Entry(email_frame, width=80)
        self.cc_email_entry.grid(row=1, column=1, padx=5, pady=5)
        self.cc_email_entry.insert(0, "k.hassan@modern-egypt.com")  # Default time or placeholder


        # Subject
        ttk.Label(email_frame, text="Subject:", font=("Helvetica", 10)).grid(row=2, column=0, padx=5, pady=5)
        self.subject_entry = ttk.Entry(email_frame, width=80)
        self.subject_entry.grid(row=2, column=1, padx=5, pady=5)
        self.subject_entry.insert(0, "Test Email Flow")  # Default time or placeholder


 


        # Body
        ttk.Label(email_frame, text="Body:", font=("Helvetica", 10)).grid(row=3, column=0)
        self.body_text = scrolledtext.ScrolledText(email_frame, height=4, width=60)
        self.body_text.grid(row=3, column=1, padx=5, pady=5)
        self.body_text.insert(tk.INSERT,
                     "- Manually insert {today} for the current date.\n"
                     "- Manually insert {now} for the current date and time.")

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

    def browse_attachments(self):
        browse_files(self.attachments, self.table_list)

    def create_schedule(self):
        schedule_frame = ttk.LabelFrame(self.main_frame, text="Schedule")
        schedule_frame.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        # Scheduled Time with hints
        ttk.Label(schedule_frame, text="Scheduled Time (HH:MM):", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.time_entry = ttk.Entry(schedule_frame, width=70)
        self.time_entry.insert(0, "08:00")  # Default time or placeholder
        self.time_entry.grid(row=0, column=1, padx=5, pady=5)

        # Weekdays
        days_frame = ttk.LabelFrame(schedule_frame, text="Weekdays", width=70)
        days_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="w")

        self.days = [tk.IntVar() for _ in range(7)]
        weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        for i, day in enumerate(weekdays):
            ttk.Checkbutton(days_frame, text=day, variable=self.days[i]).grid(row=0, column=i, padx=5, pady=5)


        
        self.selected_table_identifiers = []  # Initialize an empty list to store table identifiers

    def insert_table(self):
        selected_items = [self.table_list.get(idx) for idx in self.table_list.curselection()]
        for item in selected_items:
            if item not in self.selected_table_identifiers:
                self.selected_table_identifiers.append(item)
        insert_table_into_body(selected_items, self.body_text)

    def create_submit_button(self):
        submit_button = ttk.Button(self.main_frame, text="Submit", command=self.submit_action)
        submit_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='ew')


    def submit_action(self):
        # Now use self.selected_table_identifiers in the submit_action call
        submit_action(
            self.to_email_entry.get(),
            self.cc_email_entry.get(),
            self.subject_entry.get(),
            self.body_text.get("1.0", tk.END),
            self.attachments.get(),
            [day.get() for day in self.days],
            self.time_entry.get(),
            self.selected_table_identifiers
        )

    def insert_today(self):
        print("Inserting today's date")
        # Choose which widget to insert into, e.g., self.subject_entry or self.body_text
        self._insert_placeholder('{today}', self.body_text)  # or self.subject_entry

    def insert_now(self):
        print("Inserting current time")
        # Choose which widget to insert into
        self._insert_placeholder('{now}', self.body_text)  # or self.subject_entry

    def _insert_placeholder(self, placeholder, widget):
        print("Inserting into widget:", widget)
        widget.insert(tk.END, placeholder)

    

def launch_ui():
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()

if __name__ == "__main__":
    launch_ui()
