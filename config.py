"""
Config file of calendar-sync. Originally written by
Julian van Doorn <jvdoorn@antarc.com>.
"""

import datetime

from openpyxl.utils import column_index_from_string

SCOPES = ['https://www.googleapis.com/auth/calendar']

SCHEDULE = '/Users/julian/OneDrive - Universiteit Leiden/rooster.xlsx'
CALENDAR_ID = 'u2ljfs69g7h26gj5dufjelonn0@group.calendar.google.com'

BEGIN_TIMES_CAMPUS = ['09:00', '10:00', '11:00', '12:00', '13:30', '14:30', '15:30', '16:30', '17:30']
END_TIMES_CAMPUS = ['09:45', '10:45', '11:45', '12:45', '14:15', '15:15', '16:15', '17:15', '18:15']
BEGIN_TIMES_ONLINE = ['09:15', '10:15', '11:15', '12:15', '13:15', '14:15', '15:15', '16:15', '17:15']
END_TIMES_ONLINE = ['10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00']

FIRST_DATE = datetime.datetime(2020, 8, 31)
LAST_COLUMN = column_index_from_string('AU')
FIRST_COLUMN = column_index_from_string('C')
LAST_ROW = 56
FIRST_ROW = 3

STORAGE_FILE = 'storage'
