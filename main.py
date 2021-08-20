from arguments import parser
from config import SCHEDULE
from schedule import Schedule
from utils.cache import get_appointments_from_cache, save_appointments_to_cache
from utils.google import create_appointment, delete_remote_appointments, get_calendar_service


def main(dry: bool = False):
    calendar = get_calendar_service() if not dry else None
    schedule = Schedule(SCHEDULE)

    remote_appointments = get_appointments_from_cache()
    schedule_appointments = schedule.get_appointments_from_workbook()

    created_event_count = 0

    for appointment in schedule_appointments:
        checksum = appointment.checksum

        if checksum in remote_appointments:
            event_id, historic = remote_appointments.pop(appointment.checksum)
            appointment.remote_event_id = event_id
            continue

        if appointment.is_historic:
            continue

        if not dry:
            event_id = create_appointment(calendar, appointment)
            appointment.remote_event_id = event_id
        else:
            print(f"Would create a new event {appointment}.")
        created_event_count += 1

    events_to_be_deleted = [event_id for event_id, historic in remote_appointments.values() if not historic]
    events_to_be_saved = [appointment for appointment in schedule_appointments if
                          not appointment.is_historic and appointment.remote_event_id]

    if not dry:
        delete_remote_appointments(events_to_be_deleted, calendar)
        save_appointments_to_cache(events_to_be_saved)
    else:
        print(f"Would create {created_event_count} event(s) in total.")
        print(f"Would delete {len(events_to_be_deleted)} event(s) in total.")


if __name__ == '__main__':
    options = parser.parse_args()
    main(options.dry)
