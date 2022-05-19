from datetime import datetime

### line indexes for MS Teams *.csv reports
LINE_MEETING_START = 3
LINE_MEETING_END = 4
LINE_MEETING_ATTENDANTS = 7

### column indexes for MS Teams *.csv reports
COL_NAME = 0
COL_TOTAL_TIME = 3

### date format for MS Teams *.csv reports
DATE_FORMAT_FULL = '%d.%m.%Y, %H:%M:%S'
DATE_FORMAT_DATE = '%Y-%m-%d'

### abbreviations used in MS Teams *.csv reports for duration of attendance (currently in Polish only)
ABBR_HOUR = {'godz.'}
ABBR_MIN = {'min'}
ABBR_SEC = {'sek.'}

### column names for the output csv (currently in Polish only)
LOC_NAME = 'Osoba'
LOC_PRESENT = 'PRAWDA'
LOC_NOTPRESENT = 'FAŁSZ'

### minimum threshold for counting the meeting as attended (e.g. 0.5 means that the person must be present for at least 50% of the duration of the meeting)
THRESHOLD = 0.5

ERROR_CANNOT_PARSE_FILE = 'Could not parse file {}. This meeting will not be included in the statistics.'
ERROR_INVALID_ARGUMENT = 'Invalid argument. Perhaps you used a forbidden character in the name of the output file?'
ERROR_NO_DEFAULT_FOLDER = 'The default folder with CSV reports does not exist.'
ERROR_NO_PATH_TO_FOLDER = 'No path to folder with CSV reports. Attempting default path (.{})...'
ERROR_NO_FILENAME_IN_ARGUMENTS = 'No filename specified. Saving with default naming convention...'

FILENAME = 'attendance'
now: datetime = datetime.now()
DEFAULT_OUTPUT_FILENAME = f'{now.year}-{now.month}-{now.day}-{FILENAME}.csv'
        
CSV_DIALECT = 'excel-tab'
ENCODING = 'utf16'

DEFAULT_ATTENDANCE_LIST_PATH = '\\lists'

MANUAL = r'''
Usage:
    python attendance.py <path_to_folder> <output_file_name>

<path_to_folder> - include full or relative path; currently tested in Windows 10 only
<output_file_name> - will automatically append *.csv extension

Examples:
    attendance.py C:\Users\username\Documents\attendance_lists attendance_summary
    attendance.py attendance_lists attendance_summary'''
