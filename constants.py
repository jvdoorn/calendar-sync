from openpyxl.utils import column_index_from_string

DATE_COLUMN = column_index_from_string('B')

FIRST_SCHEDULE_ROW = 3
FIRST_SCHEDULE_COLUMN = column_index_from_string('C')
LAST_SCHEDULE_COLUMN = column_index_from_string('AU')

FIRST_DATE_CELL = 'B3'

DATE_FORMAT: str = '%Y-%m-%d'
TIME_FORMAT: str = '%H:%M:%S'
DATE_TIME_FORMAT: str = DATE_FORMAT + 'T' + TIME_FORMAT

EXCEL_DATE_FORMAT: str = '%d/%m/%y'
