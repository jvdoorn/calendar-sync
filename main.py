"""
Main file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

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

STORAGE_FILE = 'storage'


class Appointment:
    """
    Appointment class, used to store some basic information and functions to serialize and
    determine its checksum.
    """

    def __init__(self, title, appointment_type, begin_time, end_time):
        self.title = title
        self.appointment_type = appointment_type
        self.begin_time = begin_time
        self.end_time = end_time

    def checksum(self):
        """
        Generates a checksum to detect if the appointment changed.
        :return: a md5 checksum.
        """
        return hashlib.md5(
            (self.title + str(self.appointment_type) + self.begin_time + self.end_time).encode()).hexdigest()

    def serialize(self):
        """
        Serializes the appointment so we can upload it to Google.
        :return: a dictionary which can be passed to the Google API.
        """
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
    """
    Some appointment types which are used to determine par example start and
    end times.
    """
    HOLIDAY, ONLINE, CAMPUS, EXAM, EMPTY = 'HOLIDAY', 'ONLINE', 'CAMPUS', 'EXAM', 'EMPTY'


def get_appointment_type(cell):
    """
    Determines the appointment type of a cell based on its color.
    :param cell: a cell in the worksheet.
    :return: the appointment type.
    """
    if cell.value is None:
        return AppointmentType.EMPTY

    color = cell.fill.fgColor

    if str(color.rgb) == 'FF5B9BD5' or color.theme is 8:
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
    """
    Attempts to find a merged cell which contains the specified cell.
    :param cell: the cell to check.
    :param ws: the worksheet.
    :return: a merged cell range or None.
    """
    for cell_range in ws.merged_cells.ranges:
        if cell.coordinate in cell_range:
            return cell_range
    return None


def get_last_in_range(cell_range, ws):
    """
    Determines the last cell in a cell range.
    :param cell_range: a cell range.
    :param ws: the worksheet.
    :return: the last cell in the cell range.
    """
    return Cell(ws, row=cell_range.max_row, column=cell_range.max_col)


def get_next_cell(cell, ws, check_merged=True):
    """
    Determines the next cell which we need to parse.
    :param cell: the current cell.
    :param ws: the worksheet.
    :param check_merged: whether we should check if the cell is merged.
    :return: the next cell or None if this was the last cell.
    """
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
    """
    Determines the date of a cell.
    :param cell: the cell.
    :return: a String containing the date (properly formatted for Google).
    """
    return (FIRST_DATE + datetime.timedelta(
        days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)).strftime('%Y-%m-%d')


def get_begin_time(cell, appointment_type):
    """
    Determines the begin time of a cell.
    :param cell: the cell
    :param appointment_type: the appointment type of the cell.
    :return: a date and time (properly formatted for Google).
    """
    date = get_date(cell)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + BEGIN_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'
    else:
        return date + 'T' + BEGIN_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'


def get_end_time(cell, appointment_type, ws):
    """
    Determines the end time of a cell (range).
    :param cell: the start cell.
    :param appointment_type: the appointment type of the cell.
    :param ws: the worksheet.
    :return: a date and time (properly formatted for Google).
    """
    date = get_date(cell)

    cell_range = get_merged_range(cell, ws)
    if cell_range:
        cell = get_last_in_range(cell_range, ws)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'
    else:
        return date + 'T' + END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00+02:00'


def get_credentials():
    """
    Attempts to load the credentials from file or generates new ones.
    :return: the credentials to authenticate with Google.
    """
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
    """
    The main method.
    :return: nothing.
    """
    # Authenticate to Google
    credentials = get_credentials()
    # Access the Calendar API
    service = build('calendar', 'v3', credentials=credentials)

    stored_appointments = {}
    new_appointments = {}

    # Load all stored appointments
    if os.path.exists(STORAGE_FILE):
        with open(STORAGE_FILE) as f:
            for line in f:
                # Load all the checksum-id-pairs.
                (checksum, uid) = line.split()
                stored_appointments[checksum] = uid

    # Open the workbook (spreadsheet)
    workbook = load_workbook(SCHEDULE).active

    # These are the appointments that have to be created or looked up in the cache
    create_appointments = []

    # The appointments of the previous iteration
    previous_appointments = []
    # Fetch the first cell
    cell = Cell(workbook, row=FIRST_ROW, column=FIRST_COLUMN)
    while cell:
        # These are the appointments found in this cell
        appointments = []

        # Determine the title of the cell
        raw_title = workbook[cell.coordinate].value
        if raw_title:
            # Split the title if multiple are present (seperated by ' / ')
            titles = raw_title.split(' / ')

            # Determine the appointment type, begin time and end time.
            appointment_type = get_appointment_type(workbook[cell.coordinate])
            begin_time = get_begin_time(cell, appointment_type)
            end_time = get_end_time(cell, appointment_type, workbook)

            for title in titles:
                # Check if we should merge with a previous appointment
                found = False
                for i in range(len(previous_appointments)):
                    previous_appointment = previous_appointments[i]
                    if title == previous_appointment.title:
                        previous_appointment.end_time = end_time
                        appointments.append(previous_appointment)
                        previous_appointments.pop(i)
                        found = True
                        break
                # If not, create a new appointment
                if not found:
                    appointments.append(Appointment(title, appointment_type, begin_time, end_time))

            # Add any previous appointments (that we did not touch)
            create_appointments += previous_appointments
            # Determine if there is a next cell
            if get_next_cell(cell, workbook) is None:
                # If there is not this was the last iteration and we should add the current appointments
                create_appointments += appointments
            # Update the previous appointments
            previous_appointments = appointments
        else:
            # The cell was empty add the previous appointments and clear the array
            create_appointments += previous_appointments
            previous_appointments = []

        # Get the next cell
        cell = get_next_cell(cell, workbook)

    # Create all the new or updated appointments
    for appointment in create_appointments:
        # Check if we previously created the appointment
        if appointment.checksum() in stored_appointments:
            # Add it to the new cache
            new_appointments[appointment.checksum()] = stored_appointments[appointment.checksum()]
            # Remove it from the old cache (performance)
            stored_appointments.pop(appointment.checksum())
        else:
            # If we did not previously create the appointment we will do so now
            try:
                # Create the event
                event = service.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()).execute(
                    num_retries=10)
                # Add it to the new cache
                new_appointments[appointment.checksum()] = event.get('id')
                # Let the user know we created a new event
                print(f'Created event with checksum {appointment.checksum()} and id {event.get("id")}')
            except Exception as e:
                print(f'Error creating event with checksum {appointment.checksum()}: {e}.')

    # Delete any appointments that have been removed or updated in the spreadsheet.
    for uid in stored_appointments.values():
        try:
            # Delete the old event
            service.events().delete(calendarId=CALENDAR_ID, eventId=uid).execute(num_retries=10)
            # Let the user know we deleted an event
            print(f'Deleted event with uid {uid}.')
        except Exception as e:
            print(f'Error deleting event with uid {uid}: {e}.')

    # Save the updated storage to disk
    with open(STORAGE_FILE, 'w') as f:
        for checksum, uid in new_appointments.items():
            f.write(f'{checksum} {uid}\n')


if __name__ == '__main__':
    main()
