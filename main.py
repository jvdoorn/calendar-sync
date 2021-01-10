from arguments import parser
from utils import *


def main(dry: bool = False):
    calendar = get_calendar_service() if not dry else None

    remote_appointments: dict = get_stored_appointments()
    local_appointments: dict = {}
    appointments_in_workbook: list = get_appointments_from_workbook()

    for appointment in appointments_in_workbook:
        checksum = appointment.checksum()

        if checksum in remote_appointments:
            local_appointments[checksum] = remote_appointments[checksum]
            remote_appointments.pop(appointment.checksum())
        else:
            if not dry:
                event_id = create_appointment(calendar, appointment)
                if event_id is not None:
                    local_appointments[appointment.checksum()] = event_id
            else:
                print("Would create an new event " + appointment.title)

    events_to_be_deleted = list(remote_appointments.values())

    if not dry:
        delete_remote_appointments(events_to_be_deleted, calendar)
        save_appointments_to_cache(local_appointments)
    else:
        print(f"Would delete {len(events_to_be_deleted)} event(s).")


if __name__ == '__main__':
    options = parser.parse_args()
    main(options.dry)
