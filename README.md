# Productivity Tracking Tool (PTT)

PTT is a simple Python-based productivity tracking tool for Windows. It prompts users to input their current task at regular intervals and logs the data for analysis.

## Features

- Popup window prompts for task input at regular intervals.
- Logs tasks with timestamps.
- Supports tracking multiple projects and task codes.

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/orifer/PTT-Productivity-tracking-tool
   ```
   
## Usage

1. Modify `task_list.json` to define your project tasks and codes.
2. Run the `main.py` script or make the PTT tool run automatically on Windows startup (Details below):
   ```
   python main.py
   ```
3. The popup window will appear at regular intervals, prompting you to select your current task from the dropdown list and click "Confirm".

## Creating Windows Task using Task Scheduler

To make the PTT tool run automatically on Windows startup, follow these steps:

1. Open Task Scheduler.
2. Click on "Create Basic Task".
3. Set a name and description for the task.
4. Choose "When the computer starts" as the trigger.
5. Choose "Start a program" as the action.
6. Browse and select the Python executable (e.g., `C:\Users\yourUser\AppData\Local\Microsoft\WindowsApps\pythonw.exe`).
7. In the "Add arguments" field, specify the path to the `"C:\Users\yourUser\someFolder\PTT - Productivity Tracking Tool\main.py"` script.
8. Click "Finish" to create the task.

Now, the PTT tool will run automatically on Windows startup, prompting you to input your tasks at regular intervals.
