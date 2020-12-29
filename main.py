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

    for uid in remote_appointments.values():
        if not dry:
            delete_appointment(calendar, uid)
        else:
            print("Would delete an event " + uid)

    if not dry:
        with open(STORAGE_FILE, 'w') as f:
            for checksum, uid in local_appointments.items():
                f.write(f'{checksum} {uid}\n')


if __name__ == '__main__':
    options = parser.parse_args()
    main(options.dry)
