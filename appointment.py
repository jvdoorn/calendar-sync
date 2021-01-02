import datetime
import hashlib
from enum import Enum
from typing import Union

from config import CAMPUS_LOCATION, DATE_FORMAT, DATE_TIME_FORMAT, EXAMS_ALL_DAY, TIME_ZONE


class Appointment:
    def __init__(self, title: str, appointment_type: str, appointment_begin_time: datetime,
                 appointment_end_time: datetime):
        self.title = title

        self.appointment_type = appointment_type

        self.appointment_begin_time = appointment_begin_time
        self.appointment_end_time = appointment_end_time

    def checksum(self) -> str:
        data = str(self.serialize()).encode()
        return hashlib.md5(data).hexdigest()

    def _is_all_day(self) -> bool:
        return self.appointment_type == AppointmentType.HOLIDAY \
               or (EXAMS_ALL_DAY and self.appointment_type == AppointmentType.EXAM)

    def get_formatted_begin_time(self) -> dict:
        if self._is_all_day():
            return {
                'date': self.appointment_begin_time.strftime(DATE_FORMAT),
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.appointment_begin_time.strftime(DATE_TIME_FORMAT),
                'timeZone': TIME_ZONE,
            }

    def get_formatted_end_time(self) -> dict:
        if self._is_all_day():
            # Add one day - Google treats the end date as 'up to' and not as 'up to and including'. Causes issues for
            # multi-day events (such as holidays). This fix does not affect single day events.
            timestamp = self.appointment_end_time + datetime.timedelta(days=1)
            return {
                'date': timestamp.strftime(DATE_FORMAT),
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.appointment_end_time.strftime(DATE_TIME_FORMAT),
                'timeZone': TIME_ZONE,
            }

    def get_location(self) -> Union[str, None]:
        if self.appointment_type == AppointmentType.CAMPUS:
            return CAMPUS_LOCATION
        else:
            return None

    def serialize(self):
        data = {
            'summary': self.title,
            'start': self.get_formatted_begin_time(),
            'end': self.get_formatted_end_time(),
            'reminders': {
                'useDefault': True,
            }
        }

        if self.get_location():
            data['location'] = self.get_location()

        return data


class AppointmentType(Enum):
    HOLIDAY, ONLINE, CAMPUS, EXAM, EMPTY = 'HOLIDAY', 'ONLINE', 'CAMPUS', 'EXAM', 'EMPTY'
