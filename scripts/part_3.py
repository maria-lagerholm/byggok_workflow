import pandas as pd
import os
import sys
import re
from docx import Document
from docx.shared import Pt
from datetime import datetime, timedelta
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json  # Added to handle JSON operations

def find_original_info(dir_name, prop_desig_dir, info_lists, columns):
    index = prop_desig_dir.index(dir_name)
    info = {col: info_lists[col][index] for col in columns}
    if pd.notna(info['Besiktningsdag']):
        info['Besiktningsdag'] = info['Besiktningsdag'].strftime('%Y-%m-%d')
    if pd.notna(info['Klockan']):
        info['Klockan'] = info['Klockan'].strftime('%H:%M')
    return {k: str(v).strip() if pd.notna(v) else '' for k, v in info.items()}

def set_arial_11(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(11)

def should_update(current_value, original_value):
    return not current_value.strip() or current_value.strip().lower() == 'ange adress'

def process_tables(tables, original_info):
    modified = False
    for table in tables:
        for row in table.rows:
            for j, cell in enumerate(row.cells):
                for col in original_info.keys():
                    if re.search(f"{re.escape(col)}:(?!_)", cell.text.strip(), re.IGNORECASE) and j + 1 < len(row.cells):
                        next_cell = row.cells[j + 1]
                        original_value = original_info[col]
                        current_value = next_cell.text.strip()
                        if should_update(current_value, original_value):
                            next_cell.text = original_value
                            set_arial_11(next_cell)
                            modified = True
    return modified

def process_header_footer(part, original_info):
    modified = False
    for paragraph in part.paragraphs:
        for col in original_info.keys():
            match = re.search(f"{re.escape(col)}:(?!_)\s*(.*)", paragraph.text.strip(), re.IGNORECASE)
            if match:
                current_value = match.group(1).strip()
                original_value = original_info[col]
                if should_update(current_value, original_value):
                    new_text = re.sub(f"{re.escape(col)}:(?!_)\s*(.*)", f"{col}: {original_value}", paragraph.text, flags=re.IGNORECASE)
                    paragraph.text = new_text
                    set_arial_11(paragraph)
                    modified = True
    for table in part.tables:
        modified |= process_tables([table], original_info)
    return modified

def authenticate_google_calendar():
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    token_path = 'token.pickle'
    creds_path = 'credentials.json'  # Ensure this is the correct path to your credentials file

    # Check if token.pickle exists (previously saved credentials)
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for next time
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def create_google_calendar_event(service, info, created_events):
    event_name = f"Besiktning {info['Kommun']}"
    
    # Check if 'Besiktningsdag' or 'Klockan' is empty or invalid
    if not info['Klockan'] or not info['Besiktningsdag']:
        print(f"Warning: Missing 'Besiktningsdag' or 'Klockan' for {info['Adress']}. Skipping event creation.")
        return None

    try:
        # Combine Besiktningsdag and Klockan, and convert them to a datetime object
        start_time = datetime.strptime(f"{info['Besiktningsdag']} {info['Klockan'].strip()}", "%Y-%m-%d %H:%M")
    except ValueError as e:
        print(f"Error parsing date and time for {info['Adress']}: {e}")
        return None

    # Create a unique identifier for the event
    event_id = f"{event_name}_{start_time.isoformat()}_{info['Adress']}"

    # Check if the event has already been created
    if event_id in created_events:
        print(f"Event already created for {info['Adress']}. Skipping.")
        return None

    end_time = start_time + timedelta(hours=2)

    event = {
        'summary': event_name,
        'location': info['Adress'],
        'description': f"{info['Fastighetsägare']}\n{info['Telefon']}\n{info['E-post']}",
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': 'Europe/Stockholm',  # Adjust to your timezone
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'Europe/Stockholm',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 1440},  # Reminder 1 day before
            ],
        },
    }

    try:
        event_result = service.events().insert(calendarId='primary', body=event).execute()
        print(f"Event created: {event_result.get('htmlLink')}")
        created_events.append(event_id)  # Add the event ID to the list
        return event_result
    except Exception as e:
        print(f"An error occurred when creating event for {info['Adress']}: {e}")
        return None

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Load events.json to keep track of created events
    events_json_path = os.path.join(base_dir, 'events.json')
    if os.path.exists(events_json_path):
        with open(events_json_path, 'r') as f:
            created_events = json.load(f)
    else:
        created_events = []

    # Authenticate Google Calendar API
    service = authenticate_google_calendar()

    file_path = os.path.join(base_dir, 'kunder', 'kundregister.xlsx')
    print(f"Reading Excel file from: {file_path}")

    df = pd.read_excel(file_path)
    df['Besiktningsdag'] = pd.to_datetime(df['Besiktningsdag'], format='%Y-%m-%d', errors='coerce')
    df['Klockan'] = pd.to_datetime(df['Klockan'], format='%H:%M:%S', errors='coerce')

    columns = ['Adress', 'Kommun', 'Fastighetsägare', 'Uppdragsgivare', 'Postadress', 'E-post', 'Telefon',
               'Uppdragsnummer', 'Besiktningsdag', 'Klockan', 'Kostnad']
    info_lists = {col: df[col].tolist() for col in columns}
    prop_desig = df['Fastighetsbeteckning'].tolist()
    prop_desig_dir = [p.replace(':', '_').replace(' ', '_') for p in prop_desig]

    total_dirs = len(prop_desig_dir)
    processed_dirs = 0
    updated_files = 0
    created_calendar_events = 0

    for current_dir in prop_desig_dir:
        processed_dirs += 1
        dir_path = os.path.join(base_dir, 'kunder', current_dir)
        if not os.path.exists(dir_path):
            continue

        original_info = find_original_info(current_dir, prop_desig_dir, info_lists, columns)

        # Create Google Calendar Event, passing created_events
        event_result = create_google_calendar_event(service, original_info, created_events)
        if event_result:
            created_calendar_events += 1

        for file in os.listdir(dir_path):
            if file.endswith('.docx'):
                doc_path = os.path.join(dir_path, file)
                doc = Document(doc_path)
                modified = False

                for table in doc.tables:
                    modified |= process_tables([table], original_info)

                for section in doc.sections:
                    modified |= process_header_footer(section.header, original_info)
                    modified |= process_header_footer(section.first_page_header, original_info)
                    modified |= process_header_footer(section.footer, original_info)
                    modified |= process_header_footer(section.first_page_footer, original_info)

                if modified:
                    doc.save(doc_path)
                    print(f"Updated: {file} in {current_dir}")
                    updated_files += 1

        # Print progress every 10% of directories processed
        if processed_dirs % max(1, total_dirs // 10) == 0:
            print(f"Progress: {processed_dirs}/{total_dirs} directories processed")

    # Save the updated list of created events to events.json
    with open(events_json_path, 'w') as f:
        json.dump(created_events, f)

    print(f"\nProcessing complete. Updated {updated_files} files and created {created_calendar_events} calendar events across {total_dirs} directories.")

if __name__ == "__main__":
    main()