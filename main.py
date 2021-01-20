from arguments import parser
from schedule import Schedule
from utils import *


def main(dry: bool = False):
    calendar = get_calendar_service() if not dry else None
    schedule = Schedule(SCHEDULE)

    remote_appointments: dict = get_stored_appointments()
    local_appointments: dict = {}
    appointments_in_workbook: list = schedule.get_appointments_from_workbook()

    created_event_count = 0

    for appointment in appointments_in_workbook:
        checksum = appointment.checksum

        if checksum in remote_appointments:
            local_appointments[checksum] = remote_appointments[checksum]
            remote_appointments.pop(appointment.checksum)
        else:
            created_event_count += 1

            if not dry:
                event_id = create_appointment(calendar, appointment)
                if event_id is not None:
                    local_appointments[appointment.checksum] = event_id
            else:
                print(f"Would create an new event {appointment.title} ({appointment.begin_time} - {appointment.end_time}).")

    events_to_be_deleted = list(remote_appointments.values())

    if not dry:
        delete_remote_appointments(events_to_be_deleted, calendar)
        save_appointments_to_cache(local_appointments)
    else:
        print(f"Would create {created_event_count} event(s) in total.")
        print(f"Would delete {len(events_to_be_deleted)} event(s) in total.")


if __name__ == '__main__':
    options = parser.parse_args()
    main(options.dry)
