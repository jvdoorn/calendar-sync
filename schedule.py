import datetime

from openpyxl import load_workbook
from openpyxl.cell import Cell

from appointment import Appointment, AppointmentType
from config import BEGIN_TIMES_CAMPUS, BEGIN_TIMES_ONLINE, END_TIMES_CAMPUS, END_TIMES_ONLINE, FIRST_COLUMN, FIRST_DATE, \
    FIRST_ROW, \
    LAST_COLUMN, LAST_ROW
from utils import update_end_time


class Schedule:
    def __init__(self, file: str):
        self._workbook = load_workbook(file).active

    def __iter__(self):
        self._current_cell = self.get_first_cell()
        return self

    def __next__(self):
        current_cell = self._current_cell
        if current_cell is None:
            raise StopIteration
        else:
            self._current_cell = self.get_next_cell(self._current_cell)
            return current_cell

    def get_merged_range(self, cell):
        for cell_range in self._workbook.merged_cells.ranges:
            if cell.coordinate in cell_range:
                return cell_range
        return None

    def get_last_in_range(self, cell_range):
        return Cell(self._workbook, row=cell_range.max_row, column=cell_range.max_col)

    def get_next_cell(self, cell, check_merged=True):
        if cell.row == LAST_ROW and cell.column == LAST_COLUMN:
            return None

        cell_range = self.get_merged_range(cell) if check_merged else None
        if cell_range:
            return self.get_next_cell(self.get_last_in_range(cell_range), False)
        else:
            if cell.column == LAST_COLUMN:
                return Cell(self._workbook, row=cell.row + 1, column=FIRST_COLUMN)
            else:
                return Cell(self._workbook, row=cell.row, column=cell.col_idx + 1)

    def get_appointments_from_workbook(self) -> list:
        appointments_in_workbook: list = []
        previous_appointments: list = []

        for cell in iter(self):
            appointments_in_cell = []
            titles = self.get_appointment_titles_from_cell(cell)

            if len(titles) > 0:
                cell_type = self.get_cell_type(self._workbook[cell.coordinate])

                cell_begin_time = self.get_cell_begin_time(cell, cell_type)
                cell_end_time = self.get_cell_end_time(cell, cell_type)

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

        appointments_in_workbook += previous_appointments
        return appointments_in_workbook

    def get_first_cell(self):
        return Cell(self._workbook, row=FIRST_ROW, column=FIRST_COLUMN)

    def get_appointment_titles_from_cell(self, cell):
        content = self._workbook[cell.coordinate].value
        return [] if content is None else content.split(' / ')

    def get_cell_type(self, cell):
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

    def get_cell_date(self, cell):
        return FIRST_DATE + datetime.timedelta(days=(cell.row - FIRST_ROW) * 7 + (cell.column - FIRST_COLUMN) // 9)

    def get_cell_begin_time(self, cell, appointment_type):
        timestamp = self.get_cell_date(cell)

        index = (cell.column - FIRST_COLUMN) % 9
        time = BEGIN_TIMES_CAMPUS[index] if appointment_type is AppointmentType.CAMPUS else BEGIN_TIMES_ONLINE[index]

        return timestamp.replace(hour=time[0], minute=time[1])

    def get_cell_end_time(self, cell, appointment_type):
        cell_range = self.get_merged_range(cell)
        if cell_range:
            cell = self.get_last_in_range(cell_range)

        timestamp = self.get_cell_date(cell)

        index = (cell.column - FIRST_COLUMN) % 9
        time = END_TIMES_CAMPUS[index] if appointment_type is AppointmentType.CAMPUS else END_TIMES_ONLINE[index]

        return timestamp.replace(hour=time[0], minute=time[1])
