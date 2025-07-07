# EFX_lib.py

A utility Python library for automating GUI tasks on Windows, with a focus on productivity and repetitive workflows. It leverages `pyautogui`, `pygetwindow`, and other libraries to interact with the OS, automate keyboard/mouse actions, manage files, and handle application windows.

## Features

- **Screen Automation:** Locate images on the screen, click, move mouse, and check pixel colors.
- **Window Management:** Move and maximize windows, ensure applications are on the main monitor.
- **Clipboard & Typing:** Copy/paste text, simulate key presses, and control Caps Lock.
- **File Operations:** Wait for files to appear or update, check file existence, automate file saving dialogs.
- **Browser Automation:** Open URLs in Chrome and ensure browser windows are in focus.

## Requirements

- Python 3.x
- [pyautogui](https://pypi.org/project/PyAutoGUI/)
- [pygetwindow](https://pypi.org/project/PyGetWindow/)
- [pyperclip](https://pypi.org/project/pyperclip/)
- [pyscreeze](https://pypi.org/project/Pyscreeze/)

## Usage

Import the library in your Python scripts:
```python
import sys
sys.path.append("your_dir_path")
import EFX_lib as efx
```

Change the FIXED_PYG_IMG_PATH to the dir of your images to search in the screen.

Example: Locate and click an image on the screen

```python
efx.procurar('red_button.png', click=True, conf=0.95)
```

Example: Wait for a file to be updated

```python
file_last_update(r'C:\path\to\file.xlsx', time=30)
```

## Notes

- Image paths are set by default but can be customized.
- Designed for Windows automation; some features may not work on other OS.
- Many functions print status messages to the console for feedback.

## License

Proprietary or internal use. Adapt as needed for your workflow.
