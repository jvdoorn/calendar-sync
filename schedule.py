import datetime
from typing import List, Union

from openpyxl import load_workbook
from openpyxl.cell import Cell, MergedCell
from openpyxl.styles.colors import Color

from appointment import Appointment, AppointmentType
from config import BEGIN_TIMES_CAMPUS, BEGIN_TIMES_ONLINE, END_TIMES_CAMPUS, END_TIMES_ONLINE, FIRST_COLUMN, FIRST_DATE, \
    FIRST_ROW, LAST_COLUMN, LAST_ROW


class ScheduleCell:
    def __init__(self, parent_cell: Cell):
        self._parent_cell: Cell = parent_cell

        self.last_column: int = parent_cell.column
        self.last_row: int = parent_cell.row

    @property
    def first_column(self) -> int:
        return self._parent_cell.column

    @property
    def first_row(self) -> int:
        return self._parent_cell.row

    @property
    def type(self) -> AppointmentType:
        theme = self.color.theme
        rgb = str(self.color.rgb)

        if rgb == 'FF5B9BD5' or theme == 8:
            return AppointmentType.CAMPUS
        elif self.color.theme == 5:
            return AppointmentType.EXAM
        elif rgb == '00000000':
            return AppointmentType.ONLINE
        elif rgb == 'FFFFC000' or theme == 7:
            return AppointmentType.HOLIDAY
        else:
            print(f'WARNING: failed to determine appointment get_type for {self.value} with color {self.color}.')
            return AppointmentType.EMPTY

    @property
    def color(self) -> Color:
        return self._parent_cell.fill.fgColor

    @property
    def value(self):
        return self._parent_cell.value

    @property
    def titles(self):
        return [] if self.value is None else self.value.split(' / ')

    @property
    def begin_time(self) -> datetime.datetime:
        index = (self.first_column - FIRST_COLUMN) % 9

        date = FIRST_DATE + datetime.timedelta(
            days=(self.first_row - FIRST_ROW) * 7 + (self.first_column - FIRST_COLUMN) // 9)
        time = BEGIN_TIMES_CAMPUS[index] if self.type is AppointmentType.CAMPUS else BEGIN_TIMES_ONLINE[index]

        return date.replace(hour=time[0], minute=time[1])

    @property
    def end_time(self) -> datetime.datetime:
        index = (self.last_column - FIRST_COLUMN) % 9

        date = FIRST_DATE + datetime.timedelta(
            days=(self.last_row - FIRST_ROW) * 7 + (self.last_column - FIRST_COLUMN) // 9)
        time = END_TIMES_CAMPUS[index] if self.type is AppointmentType.CAMPUS else END_TIMES_ONLINE[index]

        return date.replace(hour=time[0], minute=time[1])


class Schedule:
    def __init__(self, file: str):
        self._worksheet = load_workbook(file).active

    def __iter__(self):
        self._cells = [cell for row in self.rows for cell in row]
        self._iterator_cell = ScheduleCell(self._cells[0])
        self._iterator_index = 0
        return self

    @property
    def rows(self):
        return self._worksheet.iter_rows(min_row=FIRST_ROW, max_row=LAST_ROW, min_col=FIRST_COLUMN, max_col=LAST_COLUMN)

    def __next__(self) -> ScheduleCell:
        if self._iterator_index == len(self._cells):
            raise StopIteration

        next_index = self._iterator_index + 1
        next_cell: Union[None, Cell] = None if next_index == len(self._cells) else self._cells[next_index]

        self._iterator_index = next_index

        if isinstance(next_cell, MergedCell):
            self._iterator_cell.last_column = next_cell.column
            self._iterator_cell.last_row = next_cell.row

            return self.__next__()
        else:
            current_cell = self._iterator_cell
            self._iterator_cell = None if next_cell is None else ScheduleCell(next_cell)
            return current_cell

    def get_appointments_from_workbook(self) -> List[Appointment]:
        appointments_in_workbook: List[Appointment] = []

        previous_appointments: List[Appointment] = []
        for cell in iter(self):
            current_appointments: List[Appointment] = []

            for title in cell.titles:
                appointment: Union[Appointment, None] = None

                for previous_appointment in previous_appointments:
                    if previous_appointment.title == title and previous_appointment.type == cell.type:
                        appointment = previous_appointment
                        appointment.end_time = cell.end_time

                        previous_appointments.remove(previous_appointment)
                        break

                if appointment is None:
                    appointment = Appointment(title, cell.type, cell.begin_time, cell.end_time)

                current_appointments.append(appointment)

            appointments_in_workbook += previous_appointments
            previous_appointments = current_appointments

        appointments_in_workbook += previous_appointments
        return appointments_in_workbook
