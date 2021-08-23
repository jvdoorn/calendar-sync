import datetime
import logging
import os
from typing import Dict, List, Tuple

from appointment import Appointment
from config import CACHE_FILE
from constants import DATE_FORMAT


def get_cached_remote_appointments() -> Dict[str, Tuple[str, bool]]:
    cached_appointments = {}

    if not os.path.exists(CACHE_FILE):
        logging.debug('No cache file found.')
        return cached_appointments

    with open(CACHE_FILE) as cache:
        logging.debug('Loading remote appointments from cache.')
        for line in cache:
            checksum, event_id, date = line.split()
            historic = datetime.datetime.strptime(date, DATE_FORMAT) < datetime.datetime.now()

            cached_appointments[checksum] = (event_id, historic)
    return cached_appointments


def save_remote_appointments_to_cache(appointments: List[Appointment]):
    logging.info(f'Saving {len(appointments)} to cache.')

    with open(CACHE_FILE, 'w') as cache:
        for appointment in appointments:
            cache.write(
                f'{appointment.checksum} {appointment.remote_event_id} {appointment.end_time.strftime(DATE_FORMAT)}\n')
