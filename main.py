from tkinter import ttk, messagebox
import tkinter as tk
import json
import configparser
from datetime import time, datetime
import os
import sys
from redminelib import Redmine
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Update the path to be the current directory of the script
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

# Read configuration from config.ini
config = configparser.ConfigParser()
try:
    config.read('config.ini')
except Exception as e:
    messagebox.showerror("Error", f"Failed to read configuration file: {e}")
    sys.exit(1)

# Load values from the configuration
try:
    popup_interval = int(config['General']['popup_interval'])
    start_time = time(int(config['General']['start_time']), 0)
    end_time = time(int(config['General']['end_time']), 0)
    redmine_url = config['Redmine']['url']
    api_key = config['Redmine']['api_key']
    user_id = int(config['Redmine']['user_id'])
except Exception as e:
    messagebox.showerror("Error", f"Invalid configuration: {e}")
    sys.exit(1)

# Static variables
TIME_STEP_REDMINE = popup_interval / 60 / 60
SUB_PHASE_ID = 108
ON_CALL_ID = 109
CALL_IN_ID = 110
activities_dict = {
    '-': 27,
    'AD': 8,
    'CM': 9,
    'COMM': 10,
    'CU': 11,
    'DE': 12,
    'EX': 13,
    'HW': 14,
    'IT': 15,
    'MOI': 26,
    'OP': 16,
    'PM': 17,
    'PRO': 48,
    'QA': 18,
    'RM': 118,
    'SOP': 49,
    'SP': 19,
    'SR': 20,
    'SU': 21,
    'SW': 22,
    'VL': 23,
    'VR': 24,
    'WA': 25
}

# Global variables
task_list = []
redmine_issues = {}


def save_task_list(task_list):
    try:
        with open('task_list.json', 'w') as file:
            json.dump(task_list, file, indent=4)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save task list: {e}")


def add_new_task(task_list, task_name, phase_code, task_code):
    task_list['REDMINE_TASKS'][task_name] = {
        'phase_code': phase_code,
        'task_code': task_code
    }
    save_task_list(task_list)


def check_task_list(task_name):
    # Check if its a custom task
    if task_name in task_list['CUSTOM_TASKS']:
        return task_list['CUSTOM_TASKS'][task_name]
    
    # Check if its a Redmine task
    if task_name in task_list['REDMINE_TASKS']:
        return task_list['REDMINE_TASKS'][task_name]
    else:
        def on_ok():
            try:
                phase_code = phase_code_entry.get()
                task_code = task_code_entry.get()
                add_new_task(task_list, task_name, phase_code, task_code)
                popup.destroy()
                confirm()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add new task: {e}")

        popup = tk.Toplevel()
        popup.title("Missing task info")
        tk.Label(popup, text="Please enter the phase and task codes for the new task").grid(row=0, columnspan=2, pady=(10, 10))
        
        tk.Label(popup, text="Phase Code").grid(row=1, column=0)
        tk.Label(popup, text="Task Code").grid(row=2, column=0)

        phase_code_entry = tk.Entry(popup)
        task_code_entry = tk.Entry(popup)

        phase_code_entry.grid(row=1, column=1)
        task_code_entry.grid(row=2, column=1)

        ok_button = tk.Button(popup, text="OK", command=on_ok)
        ok_button.grid(row=3, columnspan=2, pady=(10, 10))

        popup.mainloop()


def load_redmine_issues():
    try:
        redmine_issues.clear()
        assigned_issues = redmine.issue.filter(assigned_to_id=user_id)

        for issue in assigned_issues:
            redmine_issues[issue.id] = issue.subject
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Redmine issues: {e}")


def log_time_on_issue(issue_id):
    try:
        today = datetime.now().date()
        time_entries = redmine.time_entry.filter(issue_id=issue_id, spent_on=today, user_id=user_id)

        if time_entries:
            time_entry = time_entries[0]  # Assuming only one time entry per issue per day
            new_hours = time_entry.hours + TIME_STEP_REDMINE
            time_entry.hours = new_hours
            time_entry.save()
        else:
            redmine_issue = redmine_issues.get(issue_id)
            if redmine_issue:
                sub_phase_value = task_list['REDMINE_TASKS'][redmine_issue]['task_code']
                phase_code_value = task_list['REDMINE_TASKS'][redmine_issue]['phase_code']
                activity_id = activities_dict[phase_code_value]
                redmine.time_entry.create(
                    issue_id=issue_id,
                    spent_on=today,
                    hours=TIME_STEP_REDMINE,
                    activity_id=activity_id,
                    custom_fields=[
                        {'id': SUB_PHASE_ID, 'value': sub_phase_value},
                        {'id': ON_CALL_ID, 'value': ''},
                        {'id': CALL_IN_ID, 'value': ''}],
                    comments='Automated log time entry'
                )
    except Exception as e:
        messagebox.showerror("Error", f"Failed to log time on issue: {e}")


def confirm():
    try:
        now = datetime.now().strftime("%d-%m-%Y")
        task_name = task_var.get()
        task_details = check_task_list(task_name)
        phase_code = task_details["phase_code"]
        task_code = task_details["task_code"]

        # Create the structure for today
        if now not in daily_task_counts:
            daily_task_counts[now] = {
                "REDMINE": {},
                "NAVISION_TIMESHEET": {}
            }

        if task_name in daily_task_counts[now]["REDMINE"]:
            daily_task_counts[now]["REDMINE"][task_name] += TIME_STEP_REDMINE
        else:
            daily_task_counts[now]["REDMINE"][task_name] = TIME_STEP_REDMINE

        if phase_code in daily_task_counts[now]["NAVISION_TIMESHEET"]:
            if task_code in daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code]:
                daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code][task_code] += TIME_STEP_REDMINE
            else:
                daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code][task_code] = TIME_STEP_REDMINE
        else:
            daily_task_counts[now]["NAVISION_TIMESHEET"][phase_code] = {task_code: TIME_STEP_REDMINE}

        with open("daily_task_counts.json", "w", encoding='utf-8') as file:
            json.dump(daily_task_counts, file, indent=4)

        root.withdraw()

        # If its a Redmine task, then log time
        issue_id = get_key_from_value(redmine_issues, task_name)
        if (issue_id):
            log_time_on_issue(issue_id)

        json_to_excel('daily_task_counts.json', 'output.xlsx')

        root.after(popup_interval * 1000, check_and_show_popup)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to confirm task: {e}")


def get_key_from_value(my_dict, target_value):
    for key, value in my_dict.items():
        if value == target_value:
            return key
    return None


def postpone():
    root.withdraw()
    root.after(60 * 5 * 1000, check_and_show_popup)


def check_and_show_popup():
    try:
        current_time = datetime.now().time()
        if start_time <= current_time <= end_time:
            load_task_list()
            show_popup()
        else:
            root.after(popup_interval * 1000, check_and_show_popup)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to check time range: {e}")


def show_popup():
    try:
        root.geometry("400x150")
        root.deiconify()
        root.attributes("-topmost", True)
        root.update_idletasks()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = root.winfo_width()
        window_height = root.winfo_height()
        x = screen_width - window_width - 10
        y = screen_height - window_height - 10
        root.geometry(f"+{x}+{y}")

        for i in range(60):
            root.geometry(f"+{x}+{screen_height - (i * 3)}")
            root.update()
            root.after(1)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to show popup: {e}")


def load_task_list():
    global task_list

    try:
        with open("task_list.json") as file:
            task_list = json.load(file)

        # Load custom tasks from the json
        dropdown_values = list(task_list['CUSTOM_TASKS'].keys())

        # Load and append values from redmine
        load_redmine_issues()
        for key, value in redmine_issues.items():
            dropdown_values.append(value)

        # Combine both and load them into the dropdown
        task_dropdown['values'] = dropdown_values
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load task list: {e}")


def json_to_excel(json_file, excel_file):
    try:
        # Load the JSON data
        with open(json_file, 'r') as f:
            data = json.load(f)

        # Organize data by month
        monthly_data = {}
        for date_str, entries in data.items():
            date = datetime.strptime(date_str, '%d-%m-%Y')
            month_str = date.strftime('%B %Y')
            if month_str not in monthly_data:
                monthly_data[month_str] = {}
            
            navision_data = entries.get("NAVISION_TIMESHEET", {})
            for phase, tasks in navision_data.items():
                for task, hours in tasks.items():
                    if (phase, task) not in monthly_data[month_str]:
                        monthly_data[month_str][(phase, task)] = {}
                    monthly_data[month_str][(phase, task)][date.day] = hours

        # Create an Excel workbook
        wb = Workbook()
        wb.remove(wb.active)  # Remove the default sheet

        weekend_fill = PatternFill(start_color="A9A9A9", end_color="A9A9A9", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for month, tasks in monthly_data.items():
            # Prepare data for the DataFrame
            rows = []
            for (phase, task), days in tasks.items():
                row = {'PROJECT': 'SOUTHPAN', 'PHASE': phase, 'TASK': task, 'SUBPHASE': '', 'ON-CALL': '', 'CALL-IN': ''}
                row.update({day: days.get(day, None) for day in range(1, 32)})
                rows.append(row)
            
            # Create a DataFrame for the month
            df = pd.DataFrame(rows, columns=['PROJECT', 'PHASE', 'TASK', 'SUBPHASE', 'ON-CALL', 'CALL-IN'] + list(range(1, 32)))
            
            # Create a new sheet in the workbook
            ws = wb.create_sheet(title=month)
            
            # Write the DataFrame to the sheet
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    # Center align the header row
                    if r_idx == 1:
                        cell.alignment = Alignment(horizontal='center')
                    
                    # Apply number format to cells with numerical values (excluding the header row)
                    elif isinstance(value, (int, float)):
                        cell.number_format = '0.0'
                    
                    # Apply gray background to weekend columns
                    day_column = c_idx - 6  # Adjust for PROJECT, PHASE, TASK, SUBPHASE, ON-CALL, CALL-IN columns
                    if day_column > 0:
                        date_str = f"{day_column}-{month.split()[0]}-{month.split()[1]}"
                        try:
                            date = datetime.strptime(date_str, "%d-%B-%Y")
                            if date.weekday() >= 5:  # Saturday or Sunday
                                cell.fill = weekend_fill
                                if r_idx == 1:  # Apply the fill to the header
                                    cell.fill = weekend_fill
                        except ValueError:
                            continue
                    
                    # Apply the border style to all cells
                    cell.border = thin_border

            # Adjust column widths
            column_widths = {
                'A': 10,  # PROJECT
                'B': 10,  # PHASE
                'C': 10,  # TASK
                'D': 15,  # SUBPHASE
                'E': 10,  # ON-CALL
                'F': 10,  # CALL-IN
            }
            for col_idx, width in enumerate(column_widths.values(), 1):
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = width

            # Adjust the widths of the day columns
            for col in range(7, 38):  # Columns for days 1 to 31
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = 3.5

        # Save the workbook to the specified file
        wb.save(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert JSON to Excel: {e}")



try:
    redmine = Redmine(redmine_url, key=api_key)
except Exception as e:
    messagebox.showerror("Error", f"Failed to initialize Redmine client: {e}")
    sys.exit(1)

root = tk.Tk()
root.title("PTT - Productivity tracking tool v1")
root.configure(bg="#f0f0f0")

ttk.Label(root, text="What are you doing my boi?", font=("Helvetica", 14), background="#f0f0f0").pack(pady=10)
task_var = tk.StringVar()
task_dropdown = ttk.Combobox(root, textvariable=task_var, font=("Helvetica", 12), width=40, state="readonly")
task_dropdown.pack(pady=5)
load_task_list()

button_frame = ttk.Frame(root, padding=5)
button_frame.pack()

confirm_button = ttk.Button(button_frame, text="Confirm", command=confirm, style="TButton")
confirm_button.grid(row=0, column=0, padx=5)

postpone_button = ttk.Button(button_frame, text="Postpone", command=postpone, style="TButton")
postpone_button.grid(row=0, column=1, padx=5)

style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12))

root.withdraw()

try:
    with open("daily_task_counts.json") as file:
        data = file.read()
        if data:
            daily_task_counts = json.loads(data)
        else:
            daily_task_counts = {}
except FileNotFoundError:
    daily_task_counts = {}
except Exception as e:
    messagebox.showerror("Error", f"Failed to load daily task counts: {e}")
    daily_task_counts = {}

check_and_show_popup()

root.mainloop()
