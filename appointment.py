"""
Support file of calendar-sync. Contains various classes and
functions regarding appointments. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

import datetime
import hashlib
from enum import Enum
from typing import Union

from config import BEGIN_TIMES_CAMPUS, BEGIN_TIMES_ONLINE, CAMPUS_LOCATION, END_TIMES_CAMPUS, END_TIMES_ONLINE, \
    EXAMS_ALL_DAY, FIRST_COLUMN, FIRST_DATE, FIRST_ROW, TIME_ZONE
from utils import get_last_in_range, get_merged_range


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

    def checksum(self, old: bool = False) -> str:
        """
        Generates a checksum to detect if the appointment changed.
        :param old: whether or not to use the old algorithm.
        :return: a md5 checksum.
        """
        if old:
            return hashlib.md5(
                (self.title + str(self.appointment_type) + self.begin_time + self.end_time).encode()
            ).hexdigest()
        else:
            return hashlib.md5(
                str(self.serialize()).encode()
            ).hexdigest()

    def _is_all_day(self) -> bool:
        """
        Determines whether the appointment takes all day.
        :return: whether the appointment is all day.
        """
        return self.appointment_type == AppointmentType.HOLIDAY or (
                    EXAMS_ALL_DAY and self.appointment_type == AppointmentType.EXAM)

    def get_begin_time(self) -> dict:
        """
        Returns a properly formatted begin time for Google.
        :return: the begin time.
        """
        if self._is_all_day():
            return {
                'date': self.begin_time.split('T')[0],
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.begin_time,
                'timeZone': TIME_ZONE,
            }

    def get_end_time(self) -> dict:
        """
        Returns a properly formatted end time for Google.
        :return: the end time.
        """
        if self._is_all_day():
            return {
                'date': self.end_time.split('T')[0],
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.end_time,
                'timeZone': TIME_ZONE,
            }

    def get_location(self) -> Union[str, None]:
        """
        Determines the location of the appointment.
        :return: the location of the appointment as a string.
        """
        if self.appointment_type == AppointmentType.CAMPUS:
            return CAMPUS_LOCATION
        else:
            return None

    def serialize(self):
        """
        Serializes the appointment so we can upload it to Google.
        :return: a dictionary which can be passed to the Google API.
        """
        data = {
            'summary': self.title,
            'start': self.get_begin_time(),
            'end': self.get_end_time(),
            'reminders': {
                'useDefault': True,
            }
        }

        if self.get_location():
            data['location'] = self.get_location()

        return data


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


def get_begin_time(cell, appointment_type):
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
        return date + 'T' + END_TIMES_CAMPUS[(cell.column - FIRST_COLUMN) % 9] + ':00'
    else:
        return date + 'T' + END_TIMES_ONLINE[(cell.column - FIRST_COLUMN) % 9] + ':00'
