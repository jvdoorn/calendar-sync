import logging
import os
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as get_service
from googleapiclient.http import BatchHttpRequest

from config import CALENDAR_ID, SCOPES


def get_credentials():
    credentials = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return credentials


def get_calendar_service():
    credentials = get_credentials()
    return get_service('calendar', 'v3', credentials=credentials)


def delete_remote_appointments(event_ids: list, calendar):
    batch: BatchHttpRequest = calendar.new_batch_http_request()

    def callback(id, _, exception):
        if exception is not None:
            logging.exception(f'Failed to delete event: {exception}.')

    for event_id in event_ids:
        batch.add(calendar.events().delete(calendarId=CALENDAR_ID, eventId=event_id), callback)

    batch.execute()


def create_appointment(calendar, appointment):
    try:
        event = calendar.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()).execute(num_retries=10)
        logging.info(f"Created a new event {appointment}.")
        return event.get('id')
    except Exception as e:
        logging.exception(f'Error creating event {appointment}: {e}.')
        return None
