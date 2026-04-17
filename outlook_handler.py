"""outlook_handler.py
Connects to the locally running Microsoft Outlook via win32com and searches
for an email by subject, then downloads its attachment to the source folder."""

import os
import logging

logger = logging.getLogger('automation')

try:
    import win32com.client
except ImportError:
    logger.error(
        'pywin32 is not installed. Run: pip install pywin32\n'
        'Note: this module only works on Windows with Outlook \
installed and logged in.'
    )
    raise


def get_outlook_inbox(folder_name: str = 'Inbox'):
    '''
    Opens Outlook via COM and returns the specified mail folder object.

    Args:
        folder_name: The name of the Outlook folder to access
        (default: 'Inbox').

    Returns:
        An Outlook MAPIFolder object representing the requested folder.

    Raises:
        RuntimeError: If Outlook cannot be opened or the folder is not found.
    '''
    logger.info('Connecting to Microsoft Outlook...')

    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace('MAPI')
    except Exception as exc:
        raise RuntimeError(
            f'Failed to connect to Outlook. Make sure Outlook is \
installed and logged in. Error: {exc}'
        ) from exc

    # Access the default root inbox folder
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    except Exception as exc:
        raise RuntimeError(
            f'Could not access the Outlook Inbox. Error: {exc}'
        ) from exc

    if folder_name.lower() == 'inbox':
        logger.debug('Using default Inbox folder.')
        return inbox

    # Search for a named sub-folder
    for folder in inbox.Folders:
        if folder.Name.lower() == folder_name.lower():
            logger.debug('Found custom folder: %s', folder_name)
            return folder

    raise RuntimeError(
        f'Outlook folder "{folder_name}" not found inside Inbox.'
    )


def find_email_by_subject(
    folder,
    subject_keyword: str,
    max_scan: int = 50
):
    '''
    Searches for the most recent email whose subject contains
    `subject_keyword`.

    Emails are scanned from newest to oldest up to `max_scan` messages.

    Args:
        folder: An Outlook MAPIFolder object to search in.
        subject_keyword: A string that must be contained in
        the email subject (case-insensitive).
        max_scan: Maximum number of emails to inspect before giving up.

    Returns:
        The first matching Outlook MailItem object, or None if not found.
    '''
    logger.info('Searching for email with subject containing: "%s"',
                subject_keyword)

    messages = folder.Items
    messages.Sort('[ReceivedTime]', True)  # Newest first

    scanned = 0
    for message in messages:
        if scanned >= max_scan:
            break
        scanned += 1

        try:
            subject = message.Subject or ''
        except AttributeError:
            continue

        if subject_keyword.lower() in subject.lower():
            logger.info('Email found: "%s" (received: %s)',
                        subject, message.ReceivedTime)
            return message

    logger.warning(
        'No email found matching "%s" within the last %d messages.',
        subject_keyword,
        max_scan
    )
    return None


def download_attachment(email_item, destination_folder: str) -> list[str]:
    '''
    Downloads all attachments from `email_item` to `destination_folder`.

    If a file with the same name already exists, it is overwritten.

    Args:
        email_item: An Outlook MailItem object that contains attachments.
        destination_folder: Absolute path where attachments will be saved.

    Returns:
        A list of absolute file paths for each attachment that was saved.

    Raises:
        ValueError: If the email has no attachments.
        OSError: If an attachment cannot be written to disk.
    '''
    attachments = email_item.Attachments
    attachment_count = attachments.Count

    if attachment_count == 0:
        raise ValueError(
            f'The email "{email_item.Subject}" has no attachments to download.'
        )

    os.makedirs(destination_folder, exist_ok=True)
    saved_paths = []

    logger.info('Downloading %d attachment(s)...', attachment_count)

    for i in range(1, attachment_count + 1):  # Outlook COM is 1-indexed
        attachment = attachments.Item(i)
        filename = attachment.FileName
        file_path = os.path.join(destination_folder, filename)

        if os.path.exists(file_path):
            logger.debug('Attachment already exists, overwriting: %s',
                         filename)

        attachment.SaveAsFile(file_path)
        saved_paths.append(file_path)
        logger.info('Saved attachment: %s -> %s', filename, destination_folder)

    return saved_paths


def run_outlook_workflow(
    source_folder: str,
    email_subject: str,
    outlook_folder_name: str = 'Inbox',
    max_scan: int = 50
) -> list[str]:
    '''
    Full Outlook workflow: connect, search for email, download attachment(s).

    Args:
        source_folder: Destination path for the downloaded
        attachment(s) (folder X).
        email_subject: Subject keyword to search for.
        outlook_folder_name: Name of the Outlook folder to search in.
        max_scan: How many emails to scan when searching.

    Returns:
        List of absolute paths to the downloaded attachment files.

    Raises:
        RuntimeError: If Outlook cannot be accessed or no
        matching email is found.
        ValueError: If the found email has no attachments.
    '''
    inbox = get_outlook_inbox(folder_name=outlook_folder_name)
    email_item = find_email_by_subject(
        folder=inbox,
        subject_keyword=email_subject,
        max_scan=max_scan
    )

    if email_item is None:
        raise RuntimeError(
            f'Could not find an email with subject containing: \
"{email_subject}"')

    downloaded_files = download_attachment(
        email_item=email_item,
        destination_folder=source_folder
    )

    return downloaded_files
