# HardwareMonitor

HardwareMonitor is a graphical application for viewing information about a computer's hardware (CPU, GPU, motherboard, RAM, and disks) on Windows. The program uses WMI (Windows Management Instrumentation) to retrieve data.

## Features
- Displays information about the CPU (model, number of cores, clock speed, etc.)
- Displays information about the GPU (model, manufacturer, driver version, etc.)
- Retrieves data about the motherboard
- Displays detailed information about RAM
- Retrieves information about installed hard drives

## Requirements
- Windows (the program uses `win32com.client`, which works only on Windows)
- Python 3.10+
- Modules:
  - `customtkinter`
  - `pywin32`
  - `tkinter`

## Usage Example
After launching the program, a window will open displaying detailed information about your computer's hardware.

![image](https://github.com/user-attachments/assets/fd6eab67-38e2-4c46-b68c-3e4dd7b1b8eb)


## Technologies
- `customtkinter` - customized Tkinter interface
- `win32com.client` - interaction with WMI
- `tkinter` - for creating the graphical interface
- `threading` - multithreaded execution of WMI queries

## License
This project is licensed under the MIT License.
