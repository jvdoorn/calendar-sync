import datetime
from typing import List, Union

from openpyxl import load_workbook
from openpyxl.cell import Cell, MergedCell

from appointment import Appointment, AppointmentType
from config import BEGIN_TIMES_CAMPUS, BEGIN_TIMES_ONLINE, END_TIMES_CAMPUS, END_TIMES_ONLINE, FIRST_COLUMN, FIRST_DATE, \
    FIRST_ROW


class ScheduleCell:
    def __init__(self, parent_cell: Cell):
        self._parent_cell = parent_cell

        self.last_column = parent_cell.column
        self.last_row = parent_cell.row

    @property
    def type(self):
        color = self._parent_cell.fill.fgColor

        if str(self.color.rgb) == 'FF5B9BD5' or color.theme == 8:
            return AppointmentType.CAMPUS
        elif self.color.theme == 5:
            return AppointmentType.EXAM
        elif str(self.color.rgb) == '00000000':
            return AppointmentType.ONLINE
        elif str(self.color.rgb) == 'FFFFC000' or color.theme == 7:
            return AppointmentType.HOLIDAY
        else:
            print(f'WARNING: failed to determine appointment get_type for {self.value} with color {self.color}.')
            return AppointmentType.EMPTY

    @property
    def color(self):
        return self._parent_cell.fill.fgColor

    @property
    def value(self):
        return self._parent_cell.value

    @property
    def titles(self):
        return [] if self.value is None else self.value.split(' / ')

    @property
    def date(self):
        return FIRST_DATE + datetime.timedelta(
            days=(self._parent_cell.row - FIRST_ROW) * 7 + (self._parent_cell.column - FIRST_COLUMN) // 9)

    @property
    def begin_time(self):
        index = (self._parent_cell.column - FIRST_COLUMN) % 9
        time = BEGIN_TIMES_CAMPUS[index] if self.type is AppointmentType.CAMPUS else BEGIN_TIMES_ONLINE[index]

        return self.date.replace(hour=time[0], minute=time[1])

    @property
    def end_time(self):
        index = (self.last_column - FIRST_COLUMN) % 9
        time = END_TIMES_CAMPUS[index] if self.type is AppointmentType.CAMPUS else END_TIMES_ONLINE[index]

        return self.date.replace(hour=time[0], minute=time[1])


class Schedule:
    def __init__(self, file: str):
        self._worksheet = load_workbook(file).active

    def __iter__(self):
        self._cells = [cell for row in self.rows for cell in row]
        self._current_schedule_cell = ScheduleCell(self._cells[0])
        self._iterator_index = 0
        return self

    @property
    def rows(self):
        return self._worksheet.iter_rows(min_row=FIRST_ROW, min_col=FIRST_COLUMN)

    def __next__(self) -> ScheduleCell:
        if self._iterator_index == len(self._cells) - 1:
            raise StopIteration

        next_index = self._iterator_index + 1
        next_cell = self._cells[next_index]

        self._iterator_index = next_index

        if isinstance(next_cell, MergedCell):
            self._current_schedule_cell.last_column = next_cell.column
            self._current_schedule_cell.last_row = next_cell.row

            return self.__next__()
        else:
            current_cell = self._current_schedule_cell
            self._current_schedule_cell = ScheduleCell(next_cell)
            return current_cell

    def get_appointments_from_workbook(self) -> List[Appointment]:
        appointments_in_workbook: List[Appointment] = []

        previous_appointments: List[Appointment] = []
        for cell in iter(self):
            current_appointments: List[Appointment] = []

            for title in cell.titles:
                appointment: Union[Appointment, None] = None

                for previous_appointment in previous_appointments:
                    if previous_appointment.title == title and previous_appointment.appointment_type == cell.type:
                        appointment = previous_appointment
                        appointment.appointment_end_time = cell.end_time

                        previous_appointments.remove(previous_appointment)
                        break

                if appointment is None:
                    appointment = Appointment(title, cell.type, cell.begin_time, cell.end_time)

                current_appointments.append(appointment)

            appointments_in_workbook += previous_appointments
            previous_appointments = current_appointments

        appointments_in_workbook += previous_appointments
        return appointments_in_workbook