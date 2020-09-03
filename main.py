"""
Main file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

import os

from googleapiclient.discovery import build
from openpyxl import load_workbook
from openpyxl.cell import Cell

from appointment import Appointment, get_appointment_type, get_begin_time, get_end_time
from config import CALENDAR_ID, FIRST_COLUMN, FIRST_ROW, SCHEDULE, STORAGE_FILE
from utils import get_credentials, get_next_cell


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
