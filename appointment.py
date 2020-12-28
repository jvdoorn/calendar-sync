"""
Support file of calendar-sync. Contains various classes and
functions regarding appointments. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

import hashlib
from enum import Enum
from typing import Union

from config import CAMPUS_LOCATION, EXAMS_ALL_DAY, TIME_ZONE


class Appointment:
    """
    Appointment class, used to store some basic information and functions to serialize and
    determine its checksum.
    """

    def __init__(self, title, appointment_type, appointment_begin_time, appointment_end_time):
        self.title = title
        self.appointment_type = appointment_type
        self.appointment_begin_time = appointment_begin_time
        self.appointment_end_time = appointment_end_time

    def checksum(self) -> str:
        """
        Generates a checksum to detect if the appointment changed.
        :return: a md5 checksum.
        """
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
                'date': self.appointment_begin_time.split('T')[0],
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.appointment_begin_time,
                'timeZone': TIME_ZONE,
            }

    def get_end_time(self) -> dict:
        """
        Returns a properly formatted end time for Google.
        :return: the end time.
        """
        if self._is_all_day():
            return {
                'date': self.appointment_end_time.split('T')[0],
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.appointment_end_time,
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
