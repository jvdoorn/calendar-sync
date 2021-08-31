import logging
from typing import Dict, List, Tuple

from appointment import Appointment, load_appointment_meta, save_appointment_meta
from arguments import parser
from config import SCHEDULE
from schedule import Schedule
from utils.cache import get_cached_remote_appointments, save_remote_appointments_to_cache
from utils.google import create_remote_appointments, delete_remote_appointments, \
    get_calendar_service


def main(dry: bool = False):
    schedule = Schedule(SCHEDULE)

    appointment_meta = load_appointment_meta()

    remote_appointments: Dict[str, Tuple[str, bool]] = get_cached_remote_appointments()
    schedule_appointments: List[Appointment] = schedule.get_appointments_from_workbook(appointment_meta)

    appointments_to_be_created: List[Appointment] = []

    for appointment in schedule_appointments:
        try:
            event_id, _ = remote_appointments.pop(appointment.checksum)
            appointment.remote_event_id = event_id
            continue
        except KeyError:
            pass

        if not appointment.is_historic:
            appointment.meta.accessed = True
            appointments_to_be_created.append(appointment)

    appointments_to_be_deleted = [event_id for event_id, historic in remote_appointments.values() if not historic]

    if not dry:
        calendar = get_calendar_service()

        create_remote_appointments(appointments_to_be_created, calendar)
        delete_remote_appointments(appointments_to_be_deleted, calendar)

        appointments_to_be_cached = [appointment for appointment in schedule_appointments if
                                     not appointment.is_historic and appointment.remote_event_id]
        save_remote_appointments_to_cache(appointments_to_be_cached)

        save_appointment_meta(appointment_meta)
    else:
        logging.info(f"Would create {len(appointments_to_be_created)} event(s) in total.")
        logging.info(f"Would delete {len(appointments_to_be_deleted)} event(s) in total.")


if __name__ == '__main__':
    options = parser.parse_args()
    logging.basicConfig(level=logging.getLevelName(options.log_level))

    main(options.dry)
