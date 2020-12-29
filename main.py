from utils import *


def main():
    calendar = get_calendar_service()

    remote_appointments: dict = load_cached_appointments_from_disk()
    local_appointments: dict = {}
    appointments_in_workbook: list = load_appointments_from_workbook()

    for appointment in appointments_in_workbook:
        checksum = appointment.checksum()

        if checksum in remote_appointments:
            local_appointments[checksum] = remote_appointments[checksum]
            remote_appointments.pop(appointment.checksum())
        else:
            event_id = create_appointment(calendar, appointment)
            if event_id is not None:
                local_appointments[appointment.checksum()] = event_id

    for uid in remote_appointments.values():
        delete_appointment(calendar, uid)

    with open(STORAGE_FILE, 'w') as f:
        for checksum, uid in local_appointments.items():
            f.write(f'{checksum} {uid}\n')


if __name__ == '__main__':
    main()
