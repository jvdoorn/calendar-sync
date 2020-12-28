"""
Utility file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""
import os
import pickle

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as get_service
from openpyxl import load_workbook
from openpyxl.cell import Cell

from appointment import Appointment, AppointmentType
from config import *


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


def get_calendar_service():
    credentials = get_credentials()
    return get_service('calendar', 'v3', credentials=credentials)


def load_cached_appointments_from_disk():
    stored_appointments = {}

    if os.path.exists(STORAGE_FILE):
        with open(STORAGE_FILE) as f:
            for line in f:
                (checksum, uid) = line.split()
                stored_appointments[checksum] = uid

    return stored_appointments


def load_appointments_from_workbook() -> list:
    workbook = load_workbook_from_disk()
    current_cell = get_first_cell(workbook)

    appointments_in_workbook: list = []
    previous_appointments: list = []

    while current_cell:
        appointments_in_cell = []
        titles = get_appointment_titles_from_cell(current_cell, workbook)

        if len(titles) > 0:
            cell_type = get_cell_type(workbook[current_cell.coordinate])
            cell_begin_time = get_cell_begin_time(current_cell, cell_type)
            cell_end_time = get_cell_end_time(current_cell, cell_type, workbook)

            for title in titles:
                match_with_previous_appointment = False

                for i in range(len(previous_appointments)):
                    previous_appointment = previous_appointments[i]
                    if title == previous_appointment.title:
                        update_end_time(previous_appointment, cell_end_time)
                        appointments_in_cell.append(previous_appointment)
                        previous_appointments.pop(i)
                        match_with_previous_appointment = True
                        break

                if not match_with_previous_appointment:
                    new_appointment = Appointment(title, cell_type, cell_begin_time, cell_end_time)
                    appointments_in_cell.append(new_appointment)

        appointments_in_workbook += previous_appointments
        previous_appointments = appointments_in_cell

        next_cell = get_next_cell(current_cell, workbook)

        if next_cell is None:
            appointments_in_workbook += appointments_in_cell

        current_cell = next_cell

    return appointments_in_workbook


def load_workbook_from_disk():
    return load_workbook(SCHEDULE).active


def get_first_cell(workbook):
    return Cell(workbook, row=FIRST_ROW, column=FIRST_COLUMN)


def get_appointment_titles_from_cell(cell, workbook):
    content = workbook[cell.coordinate].value
    return [] if content is None else content.split(' / ')


def update_end_time(previous_appointment, new_end_time):
    previous_appointment.appointment_end_time = new_end_time


def create_appointment(calendar, appointment):
    try:
        # Create the event
        event = calendar.events().insert(calendarId=CALENDAR_ID, body=appointment.serialize()).execute(
            num_retries=10)
        # Let the user know we created a new event
        print(f'Created event with checksum {appointment.checksum()} and id {event.get("id")}')
        return event.get('id')
    except Exception as e:
        print(f'Error creating event with checksum {appointment.checksum()}: {e}.')
        return None


def delete_appointment(calendar, event_id):
    try:
        # Delete the old event
        calendar.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute(num_retries=10)
        # Let the user know we deleted an event
        print(f'Deleted event with uid {event_id}.')
    except Exception as e:
        print(f'Error deleting event with uid {event_id}: {e}.')


def get_cell_type(cell):
    """
    Determines the appointment type of a cell based on its color.
    :param cell: a cell in the worksheet.
    :return: the appointment type.
    """
    if cell.value is None:
        return AppointmentType.EMPTY

    color = cell.fill.fgColor

    if str(color.rgb) == 'FF5B9BD5' or color.theme == 8:
        return AppointmentType.CAMPUS
    elif color.theme == 5:
        return AppointmentType.EXAM
    elif str(color.rgb) == '00000000':
        return AppointmentType.ONLINE
    elif str(color.rgb) == 'FFFFC000' or color.theme == 7:
        return AppointmentType.HOLIDAY
    else:
        print(f'WARNING: failed to determine appointment type for {cell.value} with color {color}.')
        return AppointmentType.EMPTY


def get_date(cell):
    """
    Determines the date of a cell.
    :param cell: the cell.
    :return: a String containing the date (properly formatted for Google).
    """
    return (FIRST_DATE + datetime.timedelta(
        days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)).strftime('%Y-%m-%d')


def get_cell_begin_time(cell, appointment_type):
    """
    Determines the begin time of a cell.
    :param cell: the cell
    :param appointment_type: the appointment type of the cell.
    :return: a date and time (properly formatted for Google).
    """
    date = get_date(cell)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + BEGIN_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00'
    else:
        return date + 'T' + BEGIN_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00'


def get_cell_end_time(cell, appointment_type, ws):
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
        return date + 'T' + END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00'
    else:
        return date + 'T' + END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00'
