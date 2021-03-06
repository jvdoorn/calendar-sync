import datetime

from openpyxl.utils import column_index_from_string

SCOPES = ['https://www.googleapis.com/auth/calendar']

SCHEDULE = '/Users/julian/Dropbox/University/rooster.xlsx'
CALENDAR_ID = 'u2ljfs69g7h26gj5dufjelonn0@group.calendar.google.com'

BEGIN_TIMES_CAMPUS = [(9, 00), (10, 00), (11, 00), (12, 00), (12, 45), (13, 30), (14, 30), (15, 30), (16, 30)]
END_TIMES_CAMPUS = [(9, 45), (10, 45), (11, 45), (12, 45), (13, 30), (14, 15), (15, 15), (16, 15), (17, 15)]

BEGIN_TIMES_ONLINE = [(9, 15), (10, 15), (11, 15), (12, 15), (13, 15), (14, 15), (15, 15), (16, 15), (17, 15)]
END_TIMES_ONLINE = [(10, 00), (11, 00), (12, 00), (13, 00), (14, 00), (15, 00), (16, 00), (17, 00), (18, 00)]

EXAMS_ALL_DAY = True

FIRST_DATE = datetime.datetime(2020, 8, 31)
LAST_COLUMN = column_index_from_string('AU')
FIRST_COLUMN = column_index_from_string('C')
LAST_ROW = 56
FIRST_ROW = 3

CAMPUS_LOCATION = "Niels Bohrweg 2\n2333 CA Leiden\nNetherlands"

STORAGE_FILE = 'storage'

TIME_ZONE = 'Europe/Amsterdam'

DATE_FORMAT: str = '%Y-%m-%d'
TIME_FORMAT: str = '%H:%M:%S'
DATE_TIME_FORMAT: str = DATE_FORMAT + 'T' + TIME_FORMAT
