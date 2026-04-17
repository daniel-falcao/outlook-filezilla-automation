"""get_mouse_position.py
Utility helper to discover screen coordinates for use in config.py.

HOW TO USE:
  1. Run this script: python get_mouse_position.py
  2. You have 5 seconds to hover your mouse over the target UI element.
  3. The (x, y) coordinates are printed and also appended to 'coords_log.txt'.
  4. Copy the coordinates into the appropriate variable in config.py."""

import time
from datetime import datetime
import pyautogui


def capture_position(countdown_seconds: int = 5) -> tuple:
    '''
    Waits `countdown_seconds` and then captures the current mouse position.

    Args:
        countdown_seconds: Time given to position the mouse before capture.

    Returns:
        A tuple (x, y) representing the captured screen coordinates.
    '''
    print(f'Move your mouse to the target position. Capturing in {
        countdown_seconds} seconds...')

    for remaining in range(countdown_seconds, 0, -1):
        print(f'  {remaining}...', end='\r')
        time.sleep(1)

    position = pyautogui.position()
    return (position.x, position.y)


def main() -> None:
    '''
    Interactive loop that captures mouse positions one at a time and logs them.
    Press Ctrl+C to stop.
    '''
    log_file = 'coords_log.txt'
    print('=== Mouse Position Capture Utility ===')
    print(f'Captured positions will be saved to: {log_file}')
    print('Press Ctrl+C at any time to exit.\n')

    session_label = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    with open(log_file, 'a', encoding='utf-8') as log:
        log.write(f'\n--- Session: {session_label} ---\n')

        capture_index = 1
        while True:
            try:
                label = input(f'[Capture {capture_index}] Enter \
a label for this position (e.g. "Site Manager button"): ').strip()
                if not label:
                    label = f'position_{capture_index}'

                coords = capture_position(countdown_seconds=5)
                output = f'{label}: {coords}'

                print(f'\n  Captured -> {output}\n')
                log.write(output + '\n')
                log.flush()

                capture_index += 1

            except KeyboardInterrupt:
                print('\n\nExiting. Goodbye!')
                break


if __name__ == '__main__':
    main()
