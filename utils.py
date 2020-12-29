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


def get_appointments_from_workbook() -> list:
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


def get_cell_type(cell):
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


def get_cell_date(cell):
    return (FIRST_DATE + datetime.timedelta(
        days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)).strftime(DATE_FORMAT)


def get_cell_begin_time(cell, appointment_type):
    date = get_cell_date(cell)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + BEGIN_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00'
    else:
        return date + 'T' + BEGIN_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00'


def get_cell_end_time(cell, appointment_type, ws):
    date = get_cell_date(cell)

    cell_range = get_merged_range(cell, ws)
    if cell_range:
        cell = get_last_in_range(cell_range, ws)

    if appointment_type is AppointmentType.CAMPUS:
        return date + 'T' + END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00'
    else:
        return date + 'T' + END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00'
