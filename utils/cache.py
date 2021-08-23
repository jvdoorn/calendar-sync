import datetime
import os
from typing import Dict, List, Tuple

from appointment import Appointment
from config import STORAGE_FILE
from constants import DATE_FORMAT


def get_cached_remote_appointments() -> Dict[str, Tuple[str, bool]]:
    cached_appointments = {}

    if not os.path.exists(STORAGE_FILE):
        return cached_appointments

    with open(STORAGE_FILE) as cache:
        for line in cache:
            checksum, event_id, date = line.split()
            historic = datetime.datetime.strptime(date, DATE_FORMAT) < datetime.datetime.now()

            cached_appointments[checksum] = (event_id, historic)
    return cached_appointments


def save_remote_appointments_to_cache(appointments: List[Appointment]):
    with open(STORAGE_FILE, 'w') as cache:
        for appointment in appointments:
            cache.write(
                f'{appointment.checksum} {appointment.remote_event_id} {appointment.end_time.strftime(DATE_FORMAT)}\n')
