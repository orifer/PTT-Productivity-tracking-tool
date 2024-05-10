import tkinter as tk
from tkinter import ttk
import json
import datetime
import os
import sys

# Update the path to be the current directory of the script
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

# Static variables
POPUP_INTERVAL = 1800
START_TIME = datetime.time(8, 0)  # 8:00 AM
END_TIME = datetime.time(18, 0)   # 6:00 PM

# Function to handle confirm button click
def confirm():
    task_name = task_var.get()
    task_details = task_list[task_name]
    now = datetime.datetime.now().strftime("%d-%m-%Y")
    phase_code = task_details["phase_code"]
    task_code = task_details["task_code"]

    # Update daily task counts
    if now not in daily_task_counts:
        daily_task_counts[now] = {"REDMINE": {}, "NAVISION_TIMESHEET": {}}

    if task_name in daily_task_counts[now]["REDMINE"]:
        daily_task_counts[now]["REDMINE"][task_name] += .5
    else:
        daily_task_counts[now]["REDMINE"][task_name] = .5

    if phase_code in daily_task_counts[now]["NAVISION_TIMESHEET"]:
        if task_code in daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code]:
            daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code][task_code] += .5
        else:
            daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code][task_code] = .5
    else:
        daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code] = {task_code: .5}

    # Save updated daily task counts to JSON file
    with open("daily_task_counts.json", "w", encoding='utf-8') as file:
        json.dump(daily_task_counts, file, indent=4)  # Format JSON for readability
    # Hide the popup window
    hide_popup()

# Schedule the popup window to appear again after 5 minutes
def postpone():
    hide_popup()
    root.after(60 * 5 * 1000, check_and_show_popup)

# Function to hide the popup window
def hide_popup():
    root.withdraw()
    # Schedule the popup window to appear again after some time
    root.after(POPUP_INTERVAL * 1000, check_and_show_popup)

# Function to check if current time is within start and end time, and show popup if so
def check_and_show_popup():
    current_time = datetime.datetime.now().time()
    if START_TIME <= current_time <= END_TIME:
        show_popup()
    else:
        # If current time is not within the specified range, schedule to check again after some time
        root.after(POPUP_INTERVAL * 1000, check_and_show_popup)

# Function to show the popup window with animation
def show_popup():
    root.geometry("300x150")  # Set window size
    root.deiconify()
    root.attributes("-topmost", True)  # Keep window on top
    root.update_idletasks()  # Update window state
    # Calculate bottom right corner position
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    x = screen_width - window_width - 10
    y = screen_height - window_height - 10
    root.geometry(f"+{x}+{y}")  # Position window
    # Animate window deployment
    for i in range(60):
        root.geometry(f"+{x}+{screen_height - (i * 3)}")
        root.update()
        root.after(1)

# Read task list from JSON file
with open("task_list.json") as file:
    task_list = json.load(file)

# Create popup window
root = tk.Tk()
root.title("PTT - Productivity tracking tool")
root.configure(bg="#f0f0f0")

# Add elements to the window
ttk.Label(root, text="What are you doing my boi?", font=("Helvetica", 14), background="#f0f0f0").pack(pady=10)
task_var = tk.StringVar()
task_dropdown = ttk.Combobox(root, textvariable=task_var, font=("Helvetica", 12), width=25)
task_dropdown['values'] = list(task_list.keys())
task_dropdown.pack(pady=5)

# Create a frame to contain the buttons
button_frame = ttk.Frame(root, padding=5)
button_frame.pack()

# Add buttons to the frame
confirm_button = ttk.Button(button_frame, text="Confirm", command=confirm, style="TButton")
confirm_button.grid(row=0, column=0, padx=5)

postpone_button = ttk.Button(button_frame, text="Postpone", command=postpone, style="TButton")
postpone_button.grid(row=0, column=1, padx=5)

# Style for the buttons
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12))

# Initially hide the popup window
root.withdraw()

# Load daily task counts from JSON file
try:
    with open("daily_task_counts.json") as file:
        data = file.read()
        if data:
            daily_task_counts = json.loads(data)
        else:
            daily_task_counts = {}
except FileNotFoundError:
    daily_task_counts = {}

# Check if it's within the configured time range and show popup accordingly
check_and_show_popup()

root.mainloop()
