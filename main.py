import os
import pickle
from enum import Enum

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string

SCOPES = ['https://www.googleapis.com/auth/calendar']
SCHEDULE = '/Users/julian/OneDrive - Universiteit Leiden/rooster.xlsx'

BEGIN_TIMES_CAMPUS = ['09:00', '10:00', '11:00', '12:00', '13:30', '14:30', '15:30', '16:30', '17:30']
END_TIMES_CAMPUS = ['09:45', '10:45', '11:45', '12:45', '14:15', '15:15', '16:15', '17:15', '18:15']
BEGIN_TIMES_ONLINE = ['09:15', '10:15', '11:15', '12:15', '13:15', '14:15', '15:15', '16:15', '17:15']
END_TIMES_ONLINE = ['10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00']

LAST_COLUMN = column_index_from_string('AU')
FIRST_COLUMN = column_index_from_string('C')
LAST_ROW = 56
FIRST_ROW = 3


class Appointment:
    def __init__(self, title, type, begin_time, end_time):
        self.title = title
        self.type = type
        self.begin_time = begin_time
        self.end_time = end_time


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
    elif str(color.rgb) == 'FFFFC000':
        return AppointmentType.HOLIDAY

    raise Exception(f'Failed to determine AppointmentType from cell formatting {color}.')


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


def get_begin_time(cell, appointment_type):
    if appointment_type is AppointmentType.CAMPUS:
        return BEGIN_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9]
    else:
        return BEGIN_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9]


def get_end_time(cell, appointment_type, ws):
    cell_range = get_merged_range(cell, ws)
    if cell_range:
        cell = get_last_in_range(cell_range, ws)

    if appointment_type is AppointmentType.CAMPUS:
        return END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9]
    else:
        return END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9]


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
    # credentials = get_credentials()

    # service = build('calendar', 'v3', credentials=credentials)

    # Call the Calendar API
    # now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    # print('Getting the upcoming 10 events')
    # events_result = service.events().list(calendarId='primary', timeMin=now,
    #                                       maxResults=10, singleEvents=True,
    #                                       orderBy='startTime').execute()
    # events = events_result.get('items', [])
    #
    # if not events:
    #     print('No upcoming events found.')
    # for event in events:
    #     start = event['start'].get('dateTime', event['start'].get('date'))
    #     print(start, event['summary'])

    workbook = load_workbook(SCHEDULE).active
    cell = Cell(workbook, row=FIRST_ROW, column=FIRST_COLUMN)
    previous_appointments = []
    while cell:
        appointment_type = get_appointment_type(workbook[cell.coordinate])

        raw_title = workbook[cell.coordinate].value
        if raw_title:
            titles = raw_title.split(' / ')

            appointments = []
            for title in titles:
                if title is None:
                    pass

                for previous_appointment in previous_appointments:
                    if title == previous_appointment.title:
                        previous_appointment.end_time = get_end_time(cell, appointment_type, workbook)
                        appointments.append(previous_appointment)
                        break
                appointments.append(Appointment(title, appointment_type, get_begin_time(cell, appointment_type), get_end_time(cell, appointment_type, workbook)))

            previous_appointments = appointments
            print(f'{appointments[-1].begin_time} - {appointments[-1].end_time} {appointments[-1].title}')
        else:
            previous_appointments = []

        # print(f'{cell.coordinate}: {workbook[cell.coordinate].value} {get_appointment_type(workbook[cell.coordinate])}')
        cell = get_next_cell(cell, workbook)


if __name__ == '__main__':
    main()
