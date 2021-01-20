import os
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as get_service
from openpyxl import load_workbook

from config import *


def get_credentials():
    credentials = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return credentials


def get_calendar_service():
    credentials = get_credentials()
    return get_service('calendar', 'v3', credentials=credentials)


def get_stored_appointments():
    stored_appointments = {}

    if os.path.exists(STORAGE_FILE):
        with open(STORAGE_FILE) as f:
            for line in f:
                (checksum, uid) = line.split()
                stored_appointments[checksum] = uid

    return stored_appointments


def save_appointments_to_cache(appointments: dict):
    with open(STORAGE_FILE, 'w') as f:
        for checksum, uid in appointments.items():
            f.write(f'{checksum} {uid}\n')


def delete_remote_appointments(uids: list, calendar):
    for uid in uids:
        delete_appointment(calendar, uid)


def load_workbook_from_disk():
    return load_workbook(SCHEDULE).active


def update_end_time(previous_appointment, new_end_time):
    previous_appointment.end_time = new_end_time


def create_appointment(calendar, appointment):
    try:
        event = calendar.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()).execute(num_retries=10)
        print(f'Created event with checksum {appointment.checksum()} and id {event.get("id")}')
        return event.get('id')
    except Exception as e:
        print(f'Error creating event with checksum {appointment.checksum()}: {e}.')
        return None


def delete_appointment(calendar, event_id):
    try:
        calendar.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute(num_retries=10)
        print(f'Deleted event with uid {event_id}.')
    except Exception as e:
        print(f'Error deleting event with uid {event_id}: {e}.')
