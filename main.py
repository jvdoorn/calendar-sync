import datetime
import hashlib
import os
import pickle
from enum import Enum

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string

SCOPES = ['https://www.googleapis.com/auth/calendar']
SCHEDULE = '/Users/julian/OneDrive - Universiteit Leiden/rooster.xlsx'
CALENDAR_ID = 'u2ljfs69g7h26gj5dufjelonn0@group.calendar.google.com'

BEGIN_TIMES_CAMPUS = ['09:00', '10:00', '11:00', '12:00', '13:30', '14:30', '15:30', '16:30', '17:30']
END_TIMES_CAMPUS = ['09:45', '10:45', '11:45', '12:45', '14:15', '15:15', '16:15', '17:15', '18:15']
BEGIN_TIMES_ONLINE = ['09:15', '10:15', '11:15', '12:15', '13:15', '14:15', '15:15', '16:15', '17:15']
END_TIMES_ONLINE = ['10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00']

FIRST_DATE = datetime.datetime(2020, 8, 31)

LAST_COLUMN = column_index_from_string('AU')
FIRST_COLUMN = column_index_from_string('C')
LAST_ROW = 56
FIRST_ROW = 3

STORAGE_FILE = 'cache'


class Appointment:
    def __init__(self, title, appointment_type, begin_time, end_time):
        self.title = title
        self.appointment_type = appointment_type
        self.begin_time = begin_time
        self.end_time = end_time

    def checksum(self):
        return hashlib.md5(
            (self.title + str(self.appointment_type) + self.begin_time + self.end_time).encode()).hexdigest()

    def serialize(self):
        if self.appointment_type == AppointmentType.HOLIDAY:
            return {
                'summary': self.title,
                'start': {
                    'date': self.begin_time.split('T')[0],
                    'timeZone': 'Europe/Amsterdam',
                },
                'end': {
                    'date': self.end_time.split('T')[0],
                    'timeZone': 'Europe/Amsterdam',
                },
                'reminders': {
                    'useDefault': True,
                }
            }
        else:
            return {
                'summary': self.title,
                'start': {
                    'dateTime': self.begin_time,
                    'timeZone': 'Europe/Amsterdam',
                },
                'end': {
                    'dateTime': self.end_time,
                    'timeZone': 'Europe/Amsterdam',
                },
                'reminders': {
                    'useDefault': True,
                }
            }


class AppointmentType(Enum):
    HOLIDAY, ONLINE, CAMPUS, EXAM, EMPTY = 'HOLIDAY', 'ONLINE', 'CAMPUS', 'EXAM', 'EMPTY'


def get_appointment_type(cell):
    if cell.value is None:
        return AppointmentType.EMPTY

    color = cell.fill.fgColor

    if color.theme is 8:
        return AppointmentType.CAMPUS
    elif color.theme is 5:
        return AppointmentType.EXAM
    elif str(color.rgb) == '00000000':
        return AppointmentType.ONLINE
    elif str(color.rgb) == 'FFFFC000' or color.theme == 7:
        return AppointmentType.HOLIDAY
    else:
        print(f'WARNING: failed to determine appointment type for {cell.value} with color {color}.')
        return AppointmentType.EMPTY


def get_merged_range(cell, ws):
    for cell_range in ws.merged_cells.ranges:
        if cell.coordinate in cell_range:
            return cell_range
    return None


def get_last_in_range(cell_range, ws):
    return Cell(ws, row=cell_range.max_row, column=cell_range.max_col)


def get_next_cell(cell, ws, check_merged=True):
    if cell.row == LAST_ROW and cell.column == LAST_COLUMN:
        return None

    cell_range = get_merged_range(cell, ws)
    if cell_range and check_merged:
        return get_next_cell(get_last_in_range(cell_range, ws), ws, False)
    else:
        if cell.column == LAST_COLUMN:
            return Cell(ws, row=cell.row + 1, column=FIRST_COLUMN)
        else:
            return Cell(ws, row=cell.row, column=cell.col_idx + 1)


def get_date(cell):
    return (FIRST_DATE + datetime.timedelta(
        days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)).strftime('%Y-%m-%d')


def get_begin_time(cell, appointment_type):
    date = get_date(cell)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + BEGIN_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'
    else:
        return date + 'T' + BEGIN_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'


def get_end_time(cell, appointment_type, ws):
    date = get_date(cell)

    cell_range = get_merged_range(cell, ws)
    if cell_range:
        cell = get_last_in_range(cell_range, ws)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'
    else:
        return date + 'T' + END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'


def get_credentials():
    credentials = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return credentials


def main():
    credentials = get_credentials()

    service = build('calendar', 'v3', credentials=credentials)

    cache = {}
    new_cache = {}
    if os.path.exists(STORAGE_FILE):
        with open(STORAGE_FILE) as f:
            for line in f:
                (hash, id) = line.split()
                cache[hash] = id

    create_appointments = []
    workbook = load_workbook(SCHEDULE).active
    cell = Cell(workbook, row=FIRST_ROW, column=FIRST_COLUMN)
    previous_appointments = []
    while cell:
        appointment_type = get_appointment_type(workbook[cell.coordinate])
        begin_time = get_begin_time(cell, appointment_type)
        end_time = get_end_time(cell, appointment_type, workbook)

        appointments = []

        raw_title = workbook[cell.coordinate].value
        if raw_title:
            titles = raw_title.split(' / ')

            for title in titles:
                if title is None:
                    continue

                found = False
                for i in range(len(previous_appointments)):
                    previous_appointment = previous_appointments[i]
                    if title == previous_appointment.title:
                        previous_appointment.end_time = end_time
                        appointments.append(previous_appointment)
                        previous_appointments.pop(i)
                        found = True
                        break
                if not found:
                    appointments.append(Appointment(title, appointment_type, begin_time, end_time))

            create_appointments += previous_appointments
            if get_next_cell(cell, workbook) is None:
                create_appointments += appointments
            previous_appointments = appointments
        else:
            create_appointments += previous_appointments
            previous_appointments = []

        cell = get_next_cell(cell, workbook)

    for appointment in create_appointments:
        if appointment.checksum() in cache:
            # Add it to the new cache
            new_cache[appointment.checksum()] = cache[appointment.checksum()]
            # Remove it from the old cache (performance)
            cache.pop(appointment.checksum())
        else:
            try:
                # Create the event
                event = service.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()).execute(
                    num_retries=10)
                # Add it to the new cache
                new_cache[appointment.checksum()] = event.get('id')
                # Let the user know we created a new event
                print(f'Created event with checksum {appointment.checksum()} and id {event.get("id")}')
            except Exception as e:
                print(f'Error creating event with checksum {appointment.checksum()}: {e}.')

    for id in cache.values():
        try:
            # Delete the old event
            service.events().delete(calendarId=CALENDAR_ID, eventId=id).execute(num_retries=10)
            # Let the user know we deleted an event
            print(f'Deleted event with id {id}.')
        except Exception as e:
            print(f'Error deleting event with id {id}: {e}.')

    with open(STORAGE_FILE, 'w') as f:
        # Save the new cache to disk
        for hash, id in new_cache.items():
            f.write(f'{hash} {id}\n')


if __name__ == '__main__':
    main()
