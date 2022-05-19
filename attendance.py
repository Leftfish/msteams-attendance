import csv
import os
import sys

from collections import defaultdict
from datetime import datetime
from typing import DefaultDict, List, Tuple

from config import *


def get_duration_in_minutes(meeting_duration: str) -> int:
    duration: List[str] = meeting_duration.split()
    total_minutes: int = 0

    for i in range(len(duration)):
        value: str = duration[i]

        if not value.isnumeric():
            continue

        if duration[i+1] in ABBR_HOUR:
            total_minutes += 60 * int(value)
        elif duration[i+1] in ABBR_MIN:
            total_minutes += int(value)
        elif duration[i+1] in ABBR_SEC:
            total_minutes += 1/60 * int(value)

    return total_minutes


def get_attendance_filepaths(folder_path: str) -> List[str]:
    paths_to_files: List[str] = []

    for filename in os.listdir(folder_path):
        file_path: str = os.path.join(folder_path, filename)
        if os.path.isfile(file_path) and os.path.splitext(file_path)[1] == '.csv':
            paths_to_files.append(file_path)

    return paths_to_files


def get_statistics_single_meeting(csv_file: str) -> Tuple[DefaultDict[str, int], int, str]:
    with open(csv_file, mode='r', encoding=ENCODING, errors='ignore') as csvfile:

        attendance_list = csv.reader(csvfile, dialect=csv.excel_tab)
        attendance_statistics: DefaultDict[str, int] = defaultdict(int)

        for line_count, line in enumerate(attendance_list):
            if line_count == LINE_MEETING_START:
                meeting_start: datetime = datetime.strptime(line[1], DATE_FORMAT_FULL)
            elif line_count == LINE_MEETING_END:
                meeting_end: datetime = datetime.strptime(line[1], DATE_FORMAT_FULL)
            elif line_count > LINE_MEETING_ATTENDANTS:
                person_name: str = line[COL_NAME]
                time_attending: int = line[COL_TOTAL_TIME]
                attendance_statistics[person_name] += get_duration_in_minutes(time_attending)

        meeting_date = meeting_start.date().strftime(DATE_FORMAT_DATE)
        meeting_duration = int((meeting_end - meeting_start).total_seconds() / 60.0)

        return attendance_statistics, meeting_duration, meeting_date


def get_statistics_all_meetings(paths_to_files: List[str]) -> Tuple[DefaultDict[str, List[str]], List[str]]:
    attended_meetings: DefaultDict[str, List[str]] = defaultdict(list)
    meeting_dates: List[str] = []

    for file in paths_to_files:
        try:
            attendance_statistics, meeting_duration, meeting_date = get_statistics_single_meeting(file)
            meeting_dates.append(meeting_date)

            for person in attendance_statistics:
                if attendance_statistics[person] / meeting_duration > THRESHOLD:
                    attended_meetings[person].append(meeting_date)
        except:
            print(ERROR_CANNOT_PARSE_FILE.format(file))

    return attended_meetings, sorted(meeting_dates)


def sort_attendants_by_last_name(attendants: List[str]) -> List[str]:
    sorted_persons: List[List[str, str]] = []

    for person in attendants:
        split: List[str] = person.split(maxsplit=1)
        if len(split) < 2: ## if someone did not provide their name with at least two words (e.g. John.Smith instead of John Smith)
            split.append('')
        sorted_persons.append(split)

    sorted_persons.sort(key=lambda person: (person[1]))

    return sorted_persons


def save_to_csv(files: str, filename: str) -> None:
    with open(filename, 'w', newline='', encoding=ENCODING) as csvfile:
        writer = csv.writer(csvfile, dialect=CSV_DIALECT)

        field_names: List[str] = [LOC_NAME] + meeting_dates
        persons: List[str] = sort_attendants_by_last_name(attended_meetings.keys())

        writer.writerow(field_names)

        for person in persons:
            person: str = ' '.join(person).rstrip()
            row_to_write: List[str] = [person]

            for date in meeting_dates:
                row_to_write.append(LOC_PRESENT if date in attended_meetings[person] else LOC_NOTPRESENT)

            writer.writerow(row_to_write)

def show_manual_exit() -> None:
    print(MANUAL)
    sys.exit()


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'help':
        show_manual_exit()
   
    # check if the path to attendance reports was provided, if not try the default path from config.py
    try:
        attendance_lists: List[str] = get_attendance_filepaths(sys.argv[1])

    except IndexError:
        print(ERROR_NO_PATH_TO_FOLDER.format(DEFAULT_ATTENDANCE_LIST_PATH))
        
        try:
            attendance_lists: List[str] = get_attendance_filepaths(os.path.curdir + DEFAULT_ATTENDANCE_LIST_PATH)

        except FileNotFoundError:
            print(ERROR_NO_DEFAULT_FOLDER)
            show_manual_exit()
    
    # check if the output file name was provided, if not use default name from config.py
    try:
        filename: str = f'{sys.argv[2]}.csv'

    except IndexError:
        print(ERROR_NO_FILENAME_IN_ARGUMENTS)
        filename: str = DEFAULT_OUTPUT_FILENAME
      
    # get the attendance statistics...
    attended_meetings, meeting_dates = get_statistics_all_meetings(attendance_lists)
    
    # try to save them do *.csv
    try:
        save_to_csv(attendance_lists, filename)
     
    except OSError:
        print(ERROR_INVALID_ARGUMENT)
        show_manual_exit()
    
    except:
        show_manual_exit()
