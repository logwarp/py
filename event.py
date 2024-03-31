# This Python code utilizes the subprocess module to execute PowerShell commands and retrieve event logs. It then writes the output to CSV files. The script runs in an infinite loop, periodically fetching logs every 60 seconds.

import os
import csv
import time
import datetime
import subprocess

# Set script start time
script_start_time = datetime.datetime.now()

# Define log types
event_types = ['Application', 'Security', 'Setup', 'System']

# Output file folder
log_output_folder = 'C:/data/EventLogs'

# Start infinite loop
while True:
    # Loop through each log type
    for event_type in event_types:
        # Build output file path
        output_file = os.path.join(log_output_folder, f"{event_type}.csv")

        # Get latest events using PowerShell command
        powershell_command = f"Get-WinEvent -LogName {event_type} -MaxEvents 10 | Select-Object TimeCreated, Id, ProviderName, Message | ConvertTo-Csv -NoTypeInformation"
        powershell_process = subprocess.Popen(["powershell", "-Command", powershell_command], stdout=subprocess.PIPE)
        events_csv, _ = powershell_process.communicate()

        # Decode CSV output
        events_csv = events_csv.decode("utf-8").strip()

        # Export to CSV
        with open(output_file, 'a', newline='') as csvfile:
            csvfile.write(events_csv)

    # Sleep for 60 seconds
    time.sleep(60)

    # Calculate script duration
    script_end_time = datetime.datetime.now()
    script_duration = script_end_time - script_start_time
    print("Script duration:", script_duration)
