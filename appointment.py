import datetime
import hashlib
from enum import Enum
from typing import Union

from config import CAMPUS_LOCATION, EXAMS_ALL_DAY, TIME_ZONE
from constants import DATE_FORMAT, DATE_TIME_FORMAT


class AppointmentType(Enum):
    HOLIDAY, ONLINE, CAMPUS, EXAM, EMPTY = 'HOLIDAY', 'ONLINE', 'CAMPUS', 'EXAM', 'EMPTY'


class Appointment:
    def __init__(self, title: str, type: AppointmentType, begin_time: datetime, end_time: datetime):
        self.title: str = title
        self.type: AppointmentType = type

        self.begin_time: datetime = begin_time
        self.end_time: datetime = end_time

        self.remote_event_id = None

    def __str__(self):
        return f"{self.title} ({self.begin_time} - {self.end_time})"

    @property
    def checksum(self) -> str:
        data = str(self.serialize()).encode()
        return hashlib.md5(data).hexdigest()

    @property
    def is_all_day(self) -> bool:
        return self.type == AppointmentType.HOLIDAY or (EXAMS_ALL_DAY and self.type == AppointmentType.EXAM)

    @property
    def formatted_begin_time(self) -> dict:
        if self.is_all_day:
            return {
                'date': self.begin_time.strftime(DATE_FORMAT),
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.begin_time.strftime(DATE_TIME_FORMAT),
                'timeZone': TIME_ZONE,
            }

    @property
    def formatted_end_time(self) -> dict:
        if self.is_all_day:
            # Add one day - Google treats the end date as 'up to' and not as 'up to and including'. Causes issues for
            # multi-day events (such as holidays). This fix does not affect single day events.
            timestamp = self.end_time + datetime.timedelta(days=1)
            return {
                'date': timestamp.strftime(DATE_FORMAT),
                'timeZone': TIME_ZONE,
            }
        else:
            return {
                'dateTime': self.end_time.strftime(DATE_TIME_FORMAT),
                'timeZone': TIME_ZONE,
            }

    @property
    def is_historic(self) -> bool:
        return self.end_time < datetime.datetime.now()

    @property
    def location(self) -> Union[str, None]:
        if self.type == AppointmentType.CAMPUS:
            return CAMPUS_LOCATION
        else:
            return None

    def serialize(self) -> dict:
        data = {
            'summary': self.title,
            'start': self.formatted_begin_time,
            'end': self.formatted_end_time,
            'reminders': {
                'useDefault': True,
            }
        }

        if self.location:
            data['location'] = self.location

        return data
