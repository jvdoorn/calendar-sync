import datetime

from openpyxl import load_workbook
from openpyxl.cell import MergedCell

from appointment import Appointment, AppointmentType
from config import BEGIN_TIMES_CAMPUS, BEGIN_TIMES_ONLINE, END_TIMES_CAMPUS, END_TIMES_ONLINE, FIRST_COLUMN, FIRST_DATE, \
    FIRST_ROW


def get_type(cell):
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
        print(f'WARNING: failed to determine appointment get_type for {cell.value} with color {color}.')
        return AppointmentType.EMPTY


def get_date(cell):
    return FIRST_DATE + datetime.timedelta(days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)


def get_begin_time(cell):
    index = (cell.column - FIRST_COLUMN) % 9
    time = BEGIN_TIMES_CAMPUS[index] if get_type(cell) is AppointmentType.CAMPUS else BEGIN_TIMES_ONLINE[index]

    return get_date(cell).replace(hour=time[0], minute=time[1])


def get_end_time(cell, cell_type=None):
    index = (cell.column - FIRST_COLUMN) % 9
    cell_type = cell_type if cell_type else get_type(cell)
    time = END_TIMES_CAMPUS[index] if cell_type is AppointmentType.CAMPUS else END_TIMES_ONLINE[index]

    return get_date(cell).replace(hour=time[0], minute=time[1])


class Schedule:
    def __init__(self, file: str):
        self._worksheet = load_workbook(file).active

    def __iter__(self):
        rows = self._worksheet.iter_rows(min_row=FIRST_ROW, min_col=FIRST_COLUMN)
        self._cells = [cell for row in rows for cell in row]
        self._iterator_index = 0
        return self

    def __next__(self):
        if self._iterator_index == len(self._cells):
            raise StopIteration

        current_index = self._iterator_index
        self._iterator_index += 1
        return self._cells[current_index]

    def get_appointments_from_workbook(self) -> list:
        appointments_in_workbook: list = []
        previous_appointments: list = []

        for cell in iter(self):
            if isinstance(cell, MergedCell):
                for appointment in previous_appointments:
                    appointment.appointment_end_time = get_end_time(cell, appointment.appointment_type)
            else:
                appointments_in_cell = []

                for title in ([] if cell.value is None else cell.value.split(' / ')):
                    match_with_previous_appointment = False

                    for i in range(len(previous_appointments)):
                        previous_appointment = previous_appointments[i]

                        if title == previous_appointment.title:
                            previous_appointment.appointment_end_time = get_end_time(cell,
                                                                                     previous_appointment.appointment_type)
                            appointments_in_cell.append(previous_appointment)
                            previous_appointments.pop(i)
                            match_with_previous_appointment = True
                            break

                    if not match_with_previous_appointment:
                        new_appointment = Appointment(title, get_type(cell), get_begin_time(cell),
                                                      get_end_time(cell))
                        appointments_in_cell.append(new_appointment)

                appointments_in_workbook += previous_appointments
                previous_appointments = appointments_in_cell

        appointments_in_workbook += previous_appointments
        return appointments_in_workbook
