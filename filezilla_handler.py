"""filezilla_handler.py
Automates FileZilla FTP client interactions using PyAutoGUI.
Opens FileZilla, connects to a preconfigured site, navigates the remote
directory tree, and uploads a file from the local panel via
right-click menu."""

import time
import subprocess
import logging
import os

logger = logging.getLogger('automation')

try:
    import pyautogui
except ImportError:
    logger.error('pyautogui is not installed. Run: pip install pyautogui')
    raise

# Disable PyAutoGUI fail-safe pause between actions (we manage delays manually)
pyautogui.PAUSE = 0.3


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _click(position: tuple, label: str = '', double: bool = False) -> None:
    '''
    Moves the mouse to `position` and performs a single or double left-click.

    Args:
        position: (x, y) screen coordinate.
        label: Human-readable label for logging.
        double: If True, performs a double-click instead of a single-click.
    '''
    action = 'Double-clicking' if double else 'Clicking'
    logger.debug('%s %s at %s', action, label, position)
    pyautogui.moveTo(position[0], position[1], duration=0.3)
    if double:
        pyautogui.doubleClick()
    else:
        pyautogui.click()


def _right_click(position: tuple, label: str = '') -> None:
    '''
    Moves the mouse to `position` and performs a right-click.

    Args:
        position: (x, y) screen coordinate.
        label: Human-readable label for logging.
    '''
    logger.debug('Right-clicking %s at %s', label, position)
    pyautogui.moveTo(position[0], position[1], duration=0.3)
    pyautogui.rightClick()


def _wait(seconds: float, reason: str = '') -> None:
    '''
    Pauses execution for `seconds` seconds, logging the reason.

    Args:
        seconds: Number of seconds to wait.
        reason: Human-readable description of why we are waiting.
    '''
    if reason:
        logger.debug('Waiting %fs: %s', seconds, reason)
    time.sleep(seconds)


# ---------------------------------------------------------------------------
# Public workflow functions
# ---------------------------------------------------------------------------

def launch_filezilla(exe_path: str, launch_wait: float = 5) -> None:
    '''
    Launches the FileZilla application from the given executable path and
    waits for it to fully open before returning.

    Args:
        exe_path: Full path to the filezilla.exe binary.
        launch_wait: Seconds to wait after launching before interacting.

    Raises:
        FileNotFoundError: If the executable path does not
        point to a valid file.
        OSError: If the process cannot be started.
    '''
    # import os
    if not os.path.isfile(exe_path):
        raise FileNotFoundError(
            f'FileZilla executable not found at: {exe_path}\n'
            'Please update FILEZILLA_EXE_PATH in config.py.'
        )

    logger.info('Launching FileZilla from: %s', exe_path)
    with subprocess.Popen([exe_path]):
        _wait(launch_wait, 'waiting for FileZilla to fully open')


def connect_to_site(
    open_site_manager_pos: tuple,
    site_entry_pos: tuple,
    connect_button_pos: tuple,
    connect_wait: float = 5,
    action_delay: float = 1.0
) -> None:
    '''
    Opens the FileZilla Site Manager, selects the preconfigured site entry,
    and clicks the Connect button.

    Args:
        open_site_manager_pos: Screen position of the Site Manager
        toolbar icon.
        site_entry_pos: Screen position of the site name inside Site Manager.
        connect_button_pos: Screen position of the "Connect" button.
        connect_wait: Seconds to wait after clicking Connect for
        FTP to establish.
        action_delay: Short pause between UI interactions.
    '''
    logger.info('Opening FileZilla Site Manager...')
    _click(open_site_manager_pos, 'Site Manager button')
    _wait(action_delay, 'Site Manager opening')

    logger.info('Selecting configured site...')
    _click(site_entry_pos, 'site entry')
    _wait(action_delay)

    logger.info('Clicking Connect...')
    _click(connect_button_pos, 'Connect button')
    _wait(connect_wait, 'FTP connection establishing')
    logger.info('Connected to FTP site.')


def navigate_remote_directory(
    click_1_pos: tuple,
    click_2_pos: tuple,
    click_3_pos: tuple,
    action_delay: float = 1.0
) -> None:
    '''
    Navigates the remote directory tree in FileZilla's right panel by
    performing three sequential double-clicks on folder entries.

    Args:
        click_1_pos: Screen position of the first remote folder.
        click_2_pos: Screen position of the second remote folder.
        click_3_pos: Screen position of the third remote folder.
        action_delay: Pause between each folder navigation step.
    '''
    logger.info('Navigating remote directory (3 levels)...')

    _click(click_1_pos, 'remote folder level 1', double=True)
    _wait(action_delay, 'waiting for remote folder 1 to expand')

    _click(click_2_pos, 'remote folder level 2', double=True)
    _wait(action_delay, 'waiting for remote folder 2 to expand')

    _click(click_3_pos, 'remote folder level 3', double=True)
    _wait(action_delay, 'waiting for remote folder 3 to expand')

    logger.info('Remote directory navigation complete.')


def upload_local_file(
    run_number: int,
    file_pos_run_1: tuple,
    file_pos_run_2: tuple,
    upload_option_pos: tuple,
    overwrite_confirm_pos: tuple,
    action_delay: float = 1.0,
    overwrite_dialog_wait: float = 2
) -> None:
    '''
    Right-clicks the downloaded file in FileZilla's local panel and selects
    the "Upload" option. Handles the overwrite confirmation
    dialog if it appears.

    The file's screen position differs between the first and second daily run
    because the local panel list may reorder. Two separate positions are
    provided — one per run.

    Args:
        run_number: 1 for the first daily execution, 2 for the second.
        file_pos_run_1: Screen position of the file during run 1.
        file_pos_run_2: Screen position of the file during run 2.
        upload_option_pos: Screen position of the "Upload" context menu entry.
        overwrite_confirm_pos: Screen position of the overwrite
        confirmation button.
        action_delay: Short pause between UI actions.
        overwrite_dialog_wait: Seconds to wait to detect an overwrite dialog.

    Raises:
        ValueError: If `run_number` is not 1 or 2.
    '''
    if run_number not in (1, 2):
        raise ValueError(f'run_number must be 1 or 2, got: {run_number}')

    file_position = file_pos_run_1 if run_number == 1 else file_pos_run_2
    logger.info('Using file position for run #%d: %s',
                run_number, file_position)

    logger.info('Right-clicking the local file to open context menu...')
    _right_click(file_position, 'local file in FileZilla panel')
    _wait(action_delay, 'context menu opening')

    logger.info('Clicking "Upload"...')
    _click(upload_option_pos, 'Upload menu option')
    _wait(overwrite_dialog_wait, 'checking for overwrite dialog')

    # Attempt to confirm overwrite if a dialog appeared.
    # If no dialog is present, this click lands on an inert
    # area and is harmless
    # ONLY if you set overwrite_confirm_pos carefully to
    # avoid unintended clicks.
    logger.debug('Attempting to confirm overwrite dialog (if present)...')
    _click(overwrite_confirm_pos, 'overwrite confirm button (if shown)')

    logger.info('Upload action triggered successfully.')


def run_filezilla_workflow(
    exe_path: str,
    open_site_manager_pos: tuple,
    site_entry_pos: tuple,
    connect_button_pos: tuple,
    remote_dir_click_1_pos: tuple,
    remote_dir_click_2_pos: tuple,
    remote_dir_click_3_pos: tuple,
    run_number: int,
    file_pos_run_1: tuple,
    file_pos_run_2: tuple,
    upload_option_pos: tuple,
    overwrite_confirm_pos: tuple,
    launch_wait: float = 5,
    connect_wait: float = 5,
    action_delay: float = 1.0,
    overwrite_dialog_wait: float = 2
) -> None:
    '''
    Full FileZilla automation workflow: launch, connect, navigate, upload.

    Args:
        exe_path: Path to filezilla.exe.
        open_site_manager_pos: Position to open the Site Manager.
        site_entry_pos: Position of the saved site in Site Manager.
        connect_button_pos: Position of the Connect button.
        remote_dir_click_1_pos: Position of first remote folder.
        remote_dir_click_2_pos: Position of second remote folder.
        remote_dir_click_3_pos: Position of third remote folder.
        run_number: 1 or 2 to select correct local file position.
        file_pos_run_1: Local file position for run 1.
        file_pos_run_2: Local file position for run 2.
        upload_option_pos: Position of "Upload" in context menu.
        overwrite_confirm_pos: Position of overwrite confirm button.
        launch_wait: Seconds to wait after FileZilla launches.
        connect_wait: Seconds to wait after clicking Connect.
        action_delay: Short delay between UI steps.
        overwrite_dialog_wait: Wait time to detect overwrite dialog.
    '''
    launch_filezilla(exe_path=exe_path, launch_wait=launch_wait)

    connect_to_site(
        open_site_manager_pos=open_site_manager_pos,
        site_entry_pos=site_entry_pos,
        connect_button_pos=connect_button_pos,
        connect_wait=connect_wait,
        action_delay=action_delay
    )

    navigate_remote_directory(
        click_1_pos=remote_dir_click_1_pos,
        click_2_pos=remote_dir_click_2_pos,
        click_3_pos=remote_dir_click_3_pos,
        action_delay=action_delay
    )

    upload_local_file(
        run_number=run_number,
        file_pos_run_1=file_pos_run_1,
        file_pos_run_2=file_pos_run_2,
        upload_option_pos=upload_option_pos,
        overwrite_confirm_pos=overwrite_confirm_pos,
        action_delay=action_delay,
        overwrite_dialog_wait=overwrite_dialog_wait
    )
