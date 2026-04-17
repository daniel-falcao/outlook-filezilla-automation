# 📁 Outlook + FileZilla Automation

Daily Python automation that moves local files, finds a specific email in Outlook, downloads its attachment, and uploads it via FileZilla — all without manual intervention.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Requirements](#requirements)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
  - [1. Local folders](#1-local-folders)
  - [2. Email (Outlook)](#2-email-outlook)
  - [3. FileZilla](#3-filezilla)
  - [4. Mouse positions (PyAutoGUI)](#4-mouse-positions-pyautogui)
  - [5. Daily run number](#5-daily-run-number)
- [How to Find Mouse Coordinates](#how-to-find-mouse-coordinates)
- [Running the Script](#running-the-script)
- [Automatic Scheduling (Windows Task Scheduler)](#automatic-scheduling-windows-task-scheduler)
- [Execution Flow](#execution-flow)
- [Logs](#logs)
- [Troubleshooting](#troubleshooting)
- [Important Notes](#important-notes)
- [License](#license)

---

## Overview

This script is designed to run **twice a day** and performs the following steps in sequence:

1. Reads **Folder X** (source folder) and moves all existing files to **Folder Y** (destination folder).
2. Opens **Microsoft Outlook** (already logged in) and locates an email by the configured subject.
3. Downloads the email **attachment(s)** to Folder X.
4. Opens **FileZilla** and connects to a previously configured site.
5. Navigates to a remote directory via **3 levels of clicks**.
6. Right-clicks the downloaded file in the local panel and selects **Upload**.

> Duplicate files at any step are automatically **overwritten**.

---

## Requirements

| Requirement | Details |
|-------------|---------|
| Operating System | **Windows** (required for Outlook COM integration) |
| Python | **3.10+** |
| Microsoft Outlook | Installed locally with an active session |
| FileZilla | Installed, with the target site already configured and saved in Site Manager |

---

## Project Structure

```
outlook-filezilla-automation/
│
├── main.py                  # Entry point — orchestrates the entire flow
├── config.py                # ⚙️ ALL configuration lives here
├── file_manager.py          # File move operations between folders
├── outlook_handler.py       # Outlook connection and attachment download
├── filezilla_handler.py     # FileZilla automation via PyAutoGUI
├── get_mouse_position.py    # Utility to discover mouse coordinates
│
├── utils/
│   ├── __init__.py
│   └── logger.py            # Log configuration (file + console)
│
├── logs/                    # Auto-generated on first run
│   └── automation_YYYY-MM-DD.log
│
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/your-username/outlook-filezilla-automation.git
cd outlook-filezilla-automation

# 2. Create and activate a virtual environment (recommended)
python -m venv venv
venv\Scripts\activate   # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Configuration

Open the **`config.py`** file and fill in all fields marked with `'FILL_HERE'`.

### 1. Local folders

```python
SOURCE_FOLDER      = 'C:/Users/YourName/Documents/FolderX'
DESTINATION_FOLDER = 'C:/Users/YourName/Documents/FolderY'
```

### 2. Email (Outlook)

```python
EMAIL_SUBJECT      = 'Daily Sales Report'   # Text contained in the subject line
OUTLOOK_FOLDER     = 'Inbox'                # Outlook folder to search
MAX_EMAILS_TO_SCAN = 50                     # Number of emails to scan
```

### 3. FileZilla

```python
FILEZILLA_SITE_NAME = 'MyFTPServer'
FILEZILLA_EXE_PATH  = 'C:/Program Files/FileZilla FTP Client/filezilla.exe'

# The 3 remote folder names in the directory path (visual reference only)
REMOTE_DIRECTORY_CLICKS = ['public_html', 'reports', 'daily']
```

> **Important:** The site must be previously saved in FileZilla's **Site Manager** with authentication configured.

### 4. Mouse positions (PyAutoGUI)

Each `(x, y)` position represents where PyAutoGUI will move and click the mouse on screen. See the section below to find the correct values.

```python
FILEZILLA_OPEN_SITE_MANAGER_POS = (120, 45)   # Site Manager button
FILEZILLA_SITE_ENTRY_POS        = (180, 220)  # Site name in Site Manager
FILEZILLA_CONNECT_BUTTON_POS    = (380, 520)  # "Connect" button

REMOTE_DIR_CLICK_1_POS = (850, 180)  # First remote folder
REMOTE_DIR_CLICK_2_POS = (850, 210)  # Second remote folder
REMOTE_DIR_CLICK_3_POS = (850, 235)  # Third remote folder

FILE_POSITION_RUN_1 = (250, 320)  # File position on the 1st daily run
FILE_POSITION_RUN_2 = (250, 345)  # File position on the 2nd daily run

UPLOAD_OPTION_POS1    = (310, 390)  # "Upload" option in the context menu
UPLOAD_OPTION_POS2    = (310, 350)  # "Upload" option in the context menu
OVERWRITE_CONFIRM_POS = (600, 420)  # Overwrite confirmation button
```

### 5. Daily run number

Before each execution, the variable is is adjusted based on the time of execution. Change the parameters on config.py:

```python
hour_minute = datetime.now().strftime("%H:%M")
if '06:00' <= hour_minute <= '12:00':
    RUN_NUMBER = 1  # First run of the day (before noon)
else:
    RUN_NUMBER = 2  # Second run of the day (after noon)
```

> If scheduling via **Task Scheduler**, create **Two task triggers according to the configured times.**

---

## How to Find Mouse Coordinates

Use the utility included in the project:

```bash
python get_mouse_position.py
```

The script will:
1. Wait 5 seconds while you position the mouse over the desired element.
2. Capture and display the `(x, y)` coordinates.
3. Save everything to `coords_log.txt` for reference.

Repeat for each required position and copy the values into `config.py`.

---

## Running the Script

```bash
# With the virtual environment active:
python main.py
```

To test each module individually:

```bash
# Move files only
python -c "from file_manager import move_files_to_destination; move_files_to_destination('C:/FolderX', 'C:/FolderY')"

# Capture mouse position only
python get_mouse_position.py
```

---

## Automatic Scheduling (Windows Task Scheduler)

### Creating two .bat files

**`run_first.bat`** (for the 1st run):
```bat
@echo off
cd /d "C:\path\to\outlook-filezilla-automation"
call venv\Scripts\activate
python -c "import config; config.RUN_NUMBER = 1" && python main.py
```

**`run_second.bat`** (for the 2nd run):
```bat
@echo off
cd /d "C:\path\to\outlook-filezilla-automation"
call venv\Scripts\activate
python -c "import config; config.RUN_NUMBER = 2" && python main.py
```

> **Simpler alternative:** Keep `RUN_NUMBER = 1` in `config.py` and create a second copy of the project with `RUN_NUMBER = 2` in its own `config.py`.

### Configuring in Task Scheduler

1. Open **Task Scheduler** (`taskschd.msc`).
2. Click **Create Basic Task**.
3. Set the desired time (e.g., 08:00 for the 1st run).
4. Under **Action**, select "Start a program" and point to the corresponding `.bat` file.
5. Repeat for the 2nd run (e.g., 14:00).

---

## Execution Flow

```
main.py
  │
  ├─ validate_folders()          → Checks whether folders exist
  ├─ move_files_to_destination() → Moves files from X to Y (with overwrite)
  │
  ├─ run_outlook_workflow()
  │     ├─ get_outlook_inbox()       → Connects to Outlook via COM
  │     ├─ find_email_by_subject()   → Searches for email by subject
  │     └─ download_attachment()     → Saves attachment to Folder X (with overwrite)
  │
  └─ run_filezilla_workflow()
        ├─ launch_filezilla()           → Opens FileZilla.exe
        ├─ connect_to_site()            → Site Manager → Select site → Connect
        ├─ navigate_remote_directory()  → 3x double-click on remote folders
        └─ upload_local_file()          → Right-click on file → Upload → Confirm overwrite
```

---

## Logs

Logs are automatically generated in the `logs/` folder with the filename `automation_YYYY-MM-DD.log`.

Example console output:

```
[08:01:00] INFO - Starting automation — Run #1
[08:01:00] INFO - Step 1/5 — Validating folder configuration...
[08:01:01] INFO - Step 2/5 — Moving existing files from source to destination...
[08:01:01] INFO - Moved 3 file(s) to destination folder.
[08:01:02] INFO - Step 3/5 — Connecting to Outlook and searching for email...
[08:01:04] INFO - Email found: "Daily Sales Report" (received: ...)
[08:01:05] INFO - Step 4/5 — Attachment(s) downloaded: ['report.xlsx']
[08:01:05] INFO - Step 5/5 — Starting FileZilla automation...
[08:01:12] INFO - Upload action triggered successfully.
[08:01:12] INFO - Automation completed successfully — Run #1
```

---

## Troubleshooting

| Problem | Solution |
|---------|---------|
| `ModuleNotFoundError: win32com` | Run `pip install pywin32` then `python -m pywin32_postinstall -install` |
| `Failed to connect to Outlook` | Make sure Outlook is open and logged in before running the script |
| `FileZilla executable not found` | Check the path in `FILEZILLA_EXE_PATH` inside `config.py` |
| Mouse clicking in the wrong place | Use `get_mouse_position.py` to recapture coordinates |
| Email not found | Increase `MAX_EMAILS_TO_SCAN` or verify the exact subject text |
| Upload not triggered | Confirm FileZilla is in the foreground (no other windows on top) |
| Overwrite dialog does not appear | The `OVERWRITE_CONFIRM_POS` position may need adjustment — check the exact button location |

---

## Important Notes

- **Window focus:** PyAutoGUI controls the mouse globally. Avoid moving the mouse or using the computer while the script is running.
- **Screen resolution:** Mouse coordinates are specific to the resolution set at the time of capture. If the resolution changes or FileZilla is moved to another monitor, recapture all positions.
- **Security:** Do not commit `config.py` with real credentials if the repository is public. Consider using environment variables or an unversioned `.env` file.
- **Outlook COM:** The script uses the Windows COM API to access Outlook. This requires Outlook to be installed locally — it does not work with Outlook Web or Microsoft 365 via browser.

---

## License

This project is licensed under the [MIT License](LICENSE).
