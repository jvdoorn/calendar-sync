import logging
import os
import pickle
from typing import List

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as get_service
from googleapiclient.http import BatchHttpRequest

from appointment import Appointment
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

    for event_id in event_ids:
        def callback(id, response, exception):
            if exception is not None:
                logging.exception(f'Failed to delete event: {exception}.')
            else:
                logging.info(f'Deleted appointment with remote ID {event_id}.')

        batch.add(calendar.events().delete(calendarId=CALENDAR_ID, eventId=event_id), callback)

    batch.execute()


def create_remote_appointments(appointments: List[Appointment], calendar):
    batch: BatchHttpRequest = calendar.new_batch_http_request()

    for appointment in appointments:
        def callback(id, event, exception):
            if exception is not None:
                logging.exception(f'Failed to delete event: {exception}.')
            else:
                appointment.remote_event_id = event.get('id')
                logging.info(f'Created a new appointment {appointment}, with remote ID {appointment.remote_event_id}.')

        batch.add(calendar.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()), callback)

    batch.execute()
