import os
import json
from datetime import datetime
import email_flow_executor


def day_of_week_to_int(day):
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    return days.index(day)

def is_time_to_send(flow, current_time, current_day_index):
    scheduled_time = flow.get("scheduled_time", "")
    selected_days = flow.get("selected_days", [])

    is_day_to_send = selected_days[current_day_index] == 1 if selected_days else False
    is_time_to_send = current_time == scheduled_time

    return is_day_to_send and is_time_to_send

def check_and_send_emails():
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
            else:
                print(f"Not time to send for flow: {flow_file}")

if __name__ == "__main__":
    check_and_send_emails()
