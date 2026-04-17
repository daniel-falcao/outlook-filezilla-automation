"""config.py
Configuration file for the Outlook + FileZilla automation script.
Fill in the placeholders marked with 'FILL_HERE' before running the script."""

from datetime import datetime

# -----------------------------------------------------------------------------
# GENERAL SETTINGS
# -----------------------------------------------------------------------------

# Identifies which run of the day this is: 1 for first run, 2 for second run.
# Each run is started in a specific time of the day, so we can fetch the
# time of day when the script is executed and set this variable accordingly.
# This controls which mouse position is used to click the downloaded file.
hour_minute = datetime.now().strftime("%H:%M")
if 'SET TIME HERE' <= hour_minute <= 'SET TIME HERE':
    RUN_NUMBER = 1  # First run of the day (before noon)
else:
    RUN_NUMBER = 2  # Second run of the day (after noon)

# -----------------------------------------------------------------------------
# FOLDER SETTINGS
# -----------------------------------------------------------------------------

# Path to the source folder (X) where the email attachment will be downloaded
# and from where files are read before being moved.
# Example: 'C:/Users/YourName/Documents/SourceFolder'
SOURCE_FOLDER = 'FILL_HERE'

# Path to the destination folder (Y) where existing files in SOURCE_FOLDER
# will be moved to before the new download.
# Example: 'C:/Users/YourName/Documents/DestinationFolder'
DESTINATION_FOLDER = 'FILL_HERE'

# -----------------------------------------------------------------------------
# EMAIL SETTINGS (MICROSOFT OUTLOOK)
# -----------------------------------------------------------------------------

# The exact subject line (or partial subject) to search for in Outlook.
# The search is case-insensitive and checks if the subject CONTAINS this value.
# Example: 'Daily Sales Report'
EMAIL_SUBJECT = 'FILL_HERE'

# Maximum number of emails to scan when searching by subject.
# Increase if your inbox is very large and the email might be further down.
MAX_EMAILS_TO_SCAN = 50

# Folder name inside Outlook to search in. Use 'Inbox' for the default inbox.
OUTLOOK_FOLDER = 'Inbox'

# -----------------------------------------------------------------------------
# FILEZILLA SETTINGS
# -----------------------------------------------------------------------------

# The name of the site as it appears in FileZilla's Site Manager.
# This site must already be configured and saved in FileZilla before running.
# Example: 'MyFTPServer'
FILEZILLA_SITE_NAME = 'FILL_HERE'

# Full path to the FileZilla executable on this machine.
# Example: 'C:/Program Files/FileZilla FTP Client/filezilla.exe'
FILEZILLA_EXE_PATH = 'FILL_HERE'

# The remote directory to navigate to, expressed as a list of folder names
# representing each click needed (3 clicks = 3 folder levels deep).
# Example: ['public_html', 'reports', 'daily']
REMOTE_DIRECTORY_CLICKS = ['FILL_HERE', 'FILL_HERE', 'FILL_HERE']

# -----------------------------------------------------------------------------
# PYAUTOGUI MOUSE POSITION SETTINGS
# -----------------------------------------------------------------------------
# Use a helper script or Python's pyautogui.position() to discover coordinates.
# Run: python -c "import pyautogui, time; time.sleep(3);
# print(pyautogui.position())"
# then hover over the desired element within 3 seconds.

# --- FileZilla UI positions ---

# Position of the "Open Site Manager" button or menu item in FileZilla toolbar.
FILEZILLA_OPEN_SITE_MANAGER_POS = (0, 0)  # FILL_HERE: (x, y)

# Position of the site name inside the Site Manager dialog.
FILEZILLA_SITE_ENTRY_POS = (0, 0)  # FILL_HERE: (x, y)

# Position of the "Connect" button inside the Site Manager dialog.
FILEZILLA_CONNECT_BUTTON_POS = (0, 0)  # FILL_HERE: (x, y)

# Positions for the 3 remote directory folder
# clicks (right panel of FileZilla).
# Each tuple is the (x, y) coordinate of the folder to double-click.
REMOTE_DIR_CLICK_1_POS = (0, 0)  # FILL_HERE: (x, y) - first folder level
REMOTE_DIR_CLICK_2_POS = (0, 0)  # FILL_HERE: (x, y) - second folder level
REMOTE_DIR_CLICK_3_POS = (0, 0)  # FILL_HERE: (x, y) - third folder level

# --- Downloaded file positions in FileZilla local panel ---
# The position of the downloaded file in the LEFT (local) panel of FileZilla
# changes between the first and second daily run because the file list order
# may shift. Define both positions below.

# Position of the downloaded file during the FIRST daily run.
FILE_POSITION_RUN_1 = (0, 0)  # FILL_HERE: (x, y)

# Position of the downloaded file during the SECOND daily run.
FILE_POSITION_RUN_2 = (0, 0)  # FILL_HERE: (x, y)

# Position of the "Upload" option in the right-click context menu
# during the first daily run.
UPLOAD_OPTION_POS_1 = (0, 0)  # FILL_HERE: (x, y)

# Position of the "Upload" option in the right-click context menu
# during the second daily run.
UPLOAD_OPTION_POS_2 = (0, 0)  # FILL_HERE: (x, y)

# Position to click in order to confirm file overwrite (if a dialog appears).
OVERWRITE_CONFIRM_POS = (0, 0)  # FILL_HERE: (x, y)

# -----------------------------------------------------------------------------
# TIMING SETTINGS (seconds)
# -----------------------------------------------------------------------------

# How long to wait after opening FileZilla before interacting with it.
FILEZILLA_LAUNCH_WAIT = 5

# How long to wait after clicking "Connect" for the connection to establish.
FILEZILLA_CONNECT_WAIT = 5

# General short pause between PyAutoGUI actions to avoid race conditions.
ACTION_DELAY = 1.0

# How long to wait for a potential overwrite dialog to appear.
OVERWRITE_DIALOG_WAIT = 2
