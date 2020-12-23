"""
Main file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

from appointment import Appointment, get_cell_begin_time, get_cell_end_time, get_cell_type
from config import CALENDAR_ID
from utils import *

if __name__ == '__main__':
    calendar = get_calendar_service()
    workbook = load_workbook_from_disk()

    calendar_appointments: dict = load_cached_appointments_from_disk()
    workbook_appointments: dict = {}

    # These are the appointments that have to be created or looked up in the cache
    create_appointments: list = []
    # The appointments of the previous iteration
    previous_appointments: list = []

    # Fetch the first cell
    current_cell = get_first_cell(workbook)

    while current_cell:
        appointments_in_cell = []
        titles = get_appointment_titles_from_cell(current_cell, workbook)
        if len(titles) > 0:
            cell_type = get_cell_type(workbook[current_cell.coordinate])
            cell_begin_time = get_cell_begin_time(current_cell, cell_type)
            cell_end_time = get_cell_end_time(current_cell, cell_type, workbook)

            for title in titles:
                # Check if we should merge with a previous appointment
                match_with_previous_appointment = False
                for i in range(len(previous_appointments)):
                    previous_appointment = previous_appointments[i]
                    if title == previous_appointment.title:
                        update_end_time(previous_appointment, cell_end_time)
                        appointments_in_cell.append(previous_appointment)
                        previous_appointments.pop(i)
                        match_with_previous_appointment = True
                        break
                # If not, create a new appointment
                if not match_with_previous_appointment:
                    appointments_in_cell.append(Appointment(title, cell_type, cell_begin_time, cell_end_time))

            # Add any previous appointments (that we did not touch)
            create_appointments += previous_appointments
            # Determine if there is a next cell
            if get_next_cell(current_cell, workbook) is None:
                # If there is not this was the last iteration and we should add the current appointments
                create_appointments += appointments_in_cell
            # Update the previous appointments
            previous_appointments = appointments_in_cell
        else:
            # The cell was empty add the previous appointments and clear the array
            create_appointments += previous_appointments
            previous_appointments = []

        # Get the next cell
        current_cell = get_next_cell(current_cell, workbook)

    # Create all the new or updated appointments
    for appointment in create_appointments:
        # Check if we previously created the appointment
        if appointment.checksum() in calendar_appointments:
            # Add it to the new cache
            workbook_appointments[appointment.checksum()] = calendar_appointments[appointment.checksum()]
            # Remove it from the old cache (performance)
            calendar_appointments.pop(appointment.checksum())
        else:
            event_id = create_appointment(calendar, appointment)
            if event_id is not None:
                workbook_appointments[appointment.checksum()] = event_id

    # Delete any appointments that have been removed or updated in the spreadsheet.
    for uid in calendar_appointments.values():
        delete_appointment(calendar, uid)

    # Save the updated storage to disk
    with open(STORAGE_FILE, 'w') as f:
        for checksum, uid in workbook_appointments.items():
            f.write(f'{checksum} {uid}\n')
