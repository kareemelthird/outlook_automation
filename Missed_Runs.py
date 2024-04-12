import json
from datetime import datetime, timedelta
import os

def read_execution_log():
    log_file = 'email_automation.log'
    executions = {}
    try:
        with open(log_file, 'r') as file:
            for line in file:
                try:
                    # Basic parse check
                    if 'INFO' not in line and 'WARNING' not in line and 'ERROR' not in line:
                        continue  # Skip lines that don't contain log level indicators

                    # Example log format: "2024-04-12 14:50:15,675 INFO: [FlowName] Execution success"
                    parts = line.strip().split(' ')
                    if len(parts) < 6:
                        continue  # Ensure there are enough parts to avoid index errors

                    timestamp_str = parts[0] + ' ' + parts[1].split(',')[0]  # '2024-04-12 14:50:15'
                    flow_indicator = line.find('[')
                    flow_end_indicator = line.find(']')
                    if flow_indicator == -1 or flow_end_indicator == -1:
                        continue  # Skip lines without flow name indicators

                    flow_name = line[flow_indicator + 1:flow_end_indicator]  # Extract flow name

                    if flow_name not in executions:
                        executions[flow_name] = []
                    executions[flow_name].append(timestamp_str)
                except IndexError:
                    print(f"Error parsing line: {line}")  # Log problematic lines
    except Exception as e:
        print(f"Error reading execution log: {e}")
    return executions


def fetch_scheduled_flows():
    flow_folder = 'email_flows'
    scheduled_flows = {}
    try:
        for filename in os.listdir(flow_folder):
            if filename.endswith('.json'):
                with open(os.path.join(flow_folder, filename), 'r') as file:
                    flow = json.load(file)
                    flow_name = os.path.splitext(filename)[0]
                    scheduled_flows[flow_name] = {
                        'selected_days': flow['selected_days'],
                        'scheduled_time': flow['scheduled_time']
                    }
    except Exception as e:
        print(f"Error fetching scheduled flows: {e}")
    return scheduled_flows

def analyze_missed_executions():
    missed_executions = {}
    try:
        scheduled_flows = fetch_scheduled_flows()
        executions = read_execution_log()

        for flow_name, schedule in scheduled_flows.items():
            scheduled_times = [datetime.strptime(date_time, '%Y-%m-%d %H:%M:%S') for date_time in schedule['scheduled_times']]
            executed_times = [datetime.strptime(time, '%Y-%m-%d %H:%M:%S') for time in executions.get(flow_name, [])]

            for scheduled_time in scheduled_times:
                if all(scheduled_time != executed_time for executed_time in executed_times):
                    if flow_name not in missed_executions:
                        missed_executions[flow_name] = []
                    missed_executions[flow_name].append(scheduled_time.strftime('%Y-%m-%d %H:%M:%S'))
    except Exception as e:
        print(f"Error analyzing missed executions: {e}")
    return missed_executions

