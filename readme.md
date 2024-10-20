# Real Estate Inspection Workflow Automation Script

Automate customer data processing and document creation.

## What it does

- Reads customer data from an Excel file.
- Creates customer directories and updates Word templates.
- Generates Google Calendar event links.

## Platforms

- **macOS**: Run `main_script.app` (executable available).
- **Windows/Linux**: Run `main_script.py` with Python 3.x.

## Quick Start

1. **Download the Repository**

2. **For macOS Users**

   - Navigate to the `dist` directory.
   - Double-click `main_script.app` to run (it will open a separate log window, takes a few seconds to starts). You can see what it does in the log window. Close the log window to stop the script.

## Important Notes

- **Data Backup**: Backup your data before running the script. The script is still in development, so it might not work as expected.
- **Dependencies**: The `kunder` directory must be in the same folder as `scripts` directory and must contain the Excel file `kundregister.xlsx` as well as folders `mallar`.
- **Excel File**: The Excel file must contain column names exactly as they are named in the example file `kundregister.xlsx`.
---

Feel free to download and use the script. For any issues, please open an issue on this repository.
