import datetime
import hashlib
import json
import os
from enum import Enum
from typing import Dict, Optional

from config import EXAMS_ALL_DAY, META_FILE, TIME_ZONE
from constants import DATE_FORMAT, DATE_TIME_FORMAT


class AppointmentType(Enum):
    HOLIDAY, CAMPUS, EXAM, EMPTY = 'HOLIDAY', 'CAMPUS', 'EXAM', 'EMPTY'


class AppointmentMeta:
    def __init__(self, title: str, location: Optional[str] = None):
        self.title = title
        self.location = location

        self.accessed = False


def load_appointment_meta() -> Dict[str, AppointmentMeta]:
    if not os.path.exists(META_FILE):
        return {}

    with open(META_FILE) as json_file:
        data = json.load(json_file)
        return dict([
            (subject, AppointmentMeta(**meta)) for (subject, meta) in enumerate(data)
        ])


def serialize_appointment_meta(appointment_meta: Dict[str, AppointmentMeta]) -> Dict:
    return dict([
        (key, {'title': meta.title, 'location': meta.location}) for (key, meta) in appointment_meta.items() if
        meta.accessed
    ])


def save_appointment_meta(appointment_meta: Dict[str, AppointmentMeta]):
    with open(META_FILE, 'w') as json_file:
        json.dump(serialize_appointment_meta(appointment_meta), json_file, indent=4)


class Appointment:
    def __init__(self, meta: AppointmentMeta, type: AppointmentType, begin_time: datetime, end_time: datetime):
        self.meta = meta
        self.type: AppointmentType = type

        self.begin_time: datetime = begin_time
        self.end_time: datetime = end_time

        self.remote_event_id = None

    def __str__(self):
        if self.is_all_day:
            return f"{self.meta.title} ({self.begin_time.strftime(DATE_FORMAT)} - {self.end_time.strftime(DATE_FORMAT)} all day)"
        return f"{self.meta.title} ({self.begin_time} - {self.end_time})"

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
    def location(self) -> Optional[str]:
        return self.meta.location

    @property
    def title(self) -> Optional[str]:
        return self.meta.title

    def serialize(self) -> Dict:
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
