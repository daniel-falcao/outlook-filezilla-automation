""" main.py
Entry point for the Outlook + FileZilla daily automation script.

Execution flow:
  1. Read files from source folder X.
  2. Move those files to destination folder Y.
  3. Connect to Outlook and find an email by subject.
  4. Download the email attachment to folder X.
  5. Open FileZilla, connect to the configured site.
  6. Navigate to the remote directory (3 folder levels).
  7. Upload the downloaded file via right-click > Upload."""

import sys
from utils.logger import setup_logger
import config
from file_manager import validate_folders, move_files_to_destination
from outlook_handler import run_outlook_workflow
from filezilla_handler import run_filezilla_workflow


def main() -> None:
    '''
    Orchestrates the full automation pipeline.

    Steps are executed sequentially. If any step fails, the error is logged
    and the script exits with a non-zero status code.
    '''
    logger = setup_logger(log_dir='logs')
    logger.info('=' * 60)
    logger.info('Starting automation — Run #%d', config.RUN_NUMBER)
    logger.info('=' * 60)

    # ------------------------------------------------------------------
    # Step 1 & 2: Validate folders and move existing files to destination
    # ------------------------------------------------------------------
    if config.RUN_NUMBER == 1:
        try:
            logger.info('Step 1/5 — Validating folder configuration...')
            validate_folders(
                source_folder=config.SOURCE_FOLDER,
                destination_folder=config.DESTINATION_FOLDER)

            logger.info(
                'Step 2/5 — Moving existing files from source to destination.')
            moved = move_files_to_destination(
                source_folder=config.SOURCE_FOLDER,
                destination_folder=config.DESTINATION_FOLDER)
            logger.info('Step 2/5 — Moved %d file(s) to destination folder.',
                        len(moved))

        except (ValueError, FileNotFoundError, OSError) as exc:
            logger.error('File management failed: %s', exc)
            sys.exit(1)
    else:
        logger.info(
            'Step 1 & 2 skipped for Run #%d — Assuming files were already \
moved in Run #1.', config.RUN_NUMBER)

    # ------------------------------------------------------------------
    # Steps 3 & 4: Open Outlook, find email, download attachment
    # ------------------------------------------------------------------
    try:
        logger.info(
            'Step 3/5 — Connecting to Outlook and searching for email...')
        downloaded_files = run_outlook_workflow(
            source_folder=config.SOURCE_FOLDER,
            email_subject=config.EMAIL_SUBJECT,
            outlook_folder_name=config.OUTLOOK_FOLDER,
            max_scan=config.MAX_EMAILS_TO_SCAN
        )
        logger.info('Step 4/5 — Attachment(s) downloaded: %s',
                    list(downloaded_files))

    except (RuntimeError, ValueError, OSError) as exc:
        logger.error('Outlook workflow failed: %s', exc)
        sys.exit(1)

    # ------------------------------------------------------------------
    # Steps 5, 6 & 7: Launch FileZilla, connect, navigate, upload
    # ------------------------------------------------------------------
    try:
        logger.info('Step 5/5 — Starting FileZilla automation...')
        run_filezilla_workflow(
            exe_path=config.FILEZILLA_EXE_PATH,
            open_site_manager_pos=config.FILEZILLA_OPEN_SITE_MANAGER_POS,
            site_entry_pos=config.FILEZILLA_SITE_ENTRY_POS,
            connect_button_pos=config.FILEZILLA_CONNECT_BUTTON_POS,
            remote_dir_click_1_pos=config.REMOTE_DIR_CLICK_1_POS,
            remote_dir_click_2_pos=config.REMOTE_DIR_CLICK_2_POS,
            remote_dir_click_3_pos=config.REMOTE_DIR_CLICK_3_POS,
            run_number=config.RUN_NUMBER,
            file_pos_run_1=config.FILE_POSITION_RUN_1,
            file_pos_run_2=config.FILE_POSITION_RUN_2,
            upload_option_pos=(config.UPLOAD_OPTION_POS_1
                               if config.RUN_NUMBER == 1
                               else config.UPLOAD_OPTION_POS_2),
            overwrite_confirm_pos=config.OVERWRITE_CONFIRM_POS,
            launch_wait=config.FILEZILLA_LAUNCH_WAIT,
            connect_wait=config.FILEZILLA_CONNECT_WAIT,
            action_delay=config.ACTION_DELAY,
            overwrite_dialog_wait=config.OVERWRITE_DIALOG_WAIT
        )

    except (FileNotFoundError, OSError, ValueError) as exc:
        logger.error('FileZilla workflow failed: %s', exc)
        sys.exit(1)

    logger.info('=' * 60)
    logger.info('Automation completed successfully — Run #%d',
                config.RUN_NUMBER)
    logger.info('=' * 60)


if __name__ == '__main__':
    main()
