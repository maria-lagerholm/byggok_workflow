# Kundregister Automation Scripts

Automate the management of customer data, document generation, and calendar event creation for property incpections.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Dependencies](#dependencies)

## Overview

This repository contains a set of Python scripts that streamline the handling of customer data for property inspections. The scripts perform the following tasks:

1. **Directory and File Management:** Create directories based on property designations and copy relevant template files.
2. **Document Updates:** Update Word documents with accurate property information.
3. **Google Calendar Integration:** Automatically create calendar events based on inspection schedules.

## Features

- **Automated Directory Creation:** Generates directories for each property and populates them with necessary template files.
- **Dynamic Document Editing:** Updates `.docx` files with property-specific information.
- **Google Calendar Integration:** Schedules inspection events directly into Google Calendar API.
- **Progress Tracking:** Provides real-time feedback on processing progress.

## Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/kundregister-automation.git
   cd kundregister-automation
   ```

2. **Create a Virtual Environment (Optional but Recommended):**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Required Dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

## Configuration

1. **Excel File Setup:**

   - Ensure your Excel file (`kundregister.xlsx`) is placed inside the `kunder` directory.
   - The Excel file should contain the following columns:
     - `Fastighetsbeteckning`
     - `Saljare`
     - `Kopare`
     - `Besiktningsdag`
     - `Klockan`
     - `Adress`
     - `Kommun`
     - `Fastighetsägare`
     - `Uppdragsgivare`
     - `Postadress`
     - `E-post`
     - `Telefon`
     - `Uppdragsnummer`
     - `Kostnad`

2. **Template Files:**

   - Place your template `.docx` files inside the `kunder/mallar` directory.
   - Name templates with prefixes `saljare_` or `kopare_` to categorize them accordingly.

3. **Google Calendar API Setup:**

   - Obtain `credentials.json` from the [Google Cloud Console](https://console.cloud.google.com/).
   - Place `credentials.json` in the root directory of the project.
   - The script will generate `token.pickle` after the first successful authentication.

## Usage

Run the scripts sequentially to perform all tasks:

1. **Part 1: Directory and File Management**

   ```bash
   python script_part1.py
   ```

2. **Part 2: Document Updates**

   ```bash
   python script_part2.py
   ```

3. **Part 3: Google Calendar Integration**

   ```bash
   python script_part3.py
   ```

Alternatively, if all parts are combined into a single script, simply run:

```bash
python main_script.py
```

4. **For non-coding setup create an Automator.app script (MacOS) that launches the env and runs the script:**

## Project Structure

```
kundregister-automation/
│
├── kunder/
│   ├── kundregister.xlsx
│   ├── mallar/
│   │   ├── saljare_template1.docx
│   │   └── kopare_template1.docx
│   ├── Fastighetsbeteckning1/
│   │   ├── saljare_template1.docx
│   │   └── kopare_template1.docx
│   └── ...
│
├── events.json
├── credentials.json
├── token.pickle
├── requirements.txt
├── script_part1.py
├── script_part2.py
├── script_part3.py
└── README.md
```

## Dependencies

Install all dependencies using:

```bash
pip install -r requirements.txt
```

