"""file_manager.py
Handles all local file and folder operations: moving existing files from the
source folder to the destination folder, with overwrite support."""

import os
import shutil
import logging

logger = logging.getLogger('automation')


def move_files_to_destination(source_folder: str,
                              destination_folder: str) -> list[str]:
    '''
    Moves all files from `source_folder` to `destination_folder`.

    If a file with the same name already exists in the destination, it is
    overwritten. Subdirectories inside the source folder are ignored.

    Args:
        source_folder: Absolute path to the folder to
        read files from (folder X).
        destination_folder: Absolute path to the folder to
        move files into (folder Y).

    Returns:
        A list of filenames that were successfully moved.

    Raises:
        FileNotFoundError: If `source_folder` does not exist.
        OSError: If a file cannot be moved due
        to permissions or other OS errors.
    '''
    if not os.path.isdir(source_folder):
        raise FileNotFoundError(
            f'Source folder not found: {source_folder}'
        )

    os.makedirs(destination_folder, exist_ok=True)

    moved_files = []
    all_entries = os.listdir(source_folder)
    files_only = [
        entry for entry in all_entries
        if os.path.isfile(os.path.join(source_folder, entry))
    ]

    if not files_only:
        logger.info('No files found in source folder. Nothing to move.')
        return moved_files

    # logger.info(f'Found {len(files_only)} file(s) in source folder.
    # Moving to destination...')
    logger.info('Found %d file(s) in source folder. Moving to destination...',
                len(files_only))

    for filename in files_only:
        source_path = os.path.join(source_folder, filename)
        destination_path = os.path.join(destination_folder, filename)

        if os.path.exists(destination_path):
            logger.debug('File already exists in destination, overwriting: %s',
                         filename)

        shutil.move(source_path, destination_path)
        moved_files.append(filename)
        logger.info('Moved: %s -> %s', filename, destination_folder)

    return moved_files


def validate_folders(source_folder: str, destination_folder: str) -> None:
    '''
    Validates that the source folder path is defined and the destination
    folder path is defined. Creates destination if it does not exist yet.

    Args:
        source_folder: Path to source folder X.
        destination_folder: Path to destination folder Y.

    Raises:
        ValueError: If any path is still set to the placeholder value.
        FileNotFoundError: If source_folder does not exist on disk.
    '''
    placeholder = 'FILL_HERE'

    if source_folder == placeholder or not source_folder:
        raise ValueError(
            'SOURCE_FOLDER is not configured. '
            'Please edit config.py and set the correct path.'
        )

    if destination_folder == placeholder or not destination_folder:
        raise ValueError(
            'DESTINATION_FOLDER is not configured. '
            'Please edit config.py and set the correct path.'
        )

    if not os.path.isdir(source_folder):
        raise FileNotFoundError(
            f'Source folder does not exist on disk: {source_folder}'
        )

    os.makedirs(destination_folder, exist_ok=True)
    logger.debug('Folder paths validated successfully.')
