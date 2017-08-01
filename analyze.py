from __future__ import print_function

import os
import csv
import sys
from collections import namedtuple, defaultdict

Event = namedtuple('Event', ['user', 'start', 'end', 'duration_hours', 'summary', 'creator', 'attendees'])

def get_files(file_dir):
    return os.listdir(file_dir)

def read_file(file):
    user = file.split('-')[-1].replace('.csv', '')
    with open(file, 'r') as f:
        c = csv.reader(f)
        c.next() # Skip header
        return [Event(user, *row) for row in c]
    print('Failed to retrieve events from', file)
    return []

def filter(events):
    # Remove events > 24 hours (vacations) and no attendees (todos)
    return [e for e in events if float(e.duration_hours) < 24 and len(e.attendees) > 0]

def analyze(events, username = None):
    total_hours = 0.0
    by_month_hours = defaultdict(float)
    by_month_attendees = defaultdict(int)
    by_month_events = defaultdict(int)

    for e in events:
        month = e.start[:7]
        duration_hours = float(e.duration_hours)
        total_hours += duration_hours
        by_month_hours[month] += duration_hours
        by_month_events[month] += 1
        if len(e.attendees):
            by_month_attendees[month] += len(e.attendees.split('|'))

    if len(events):
        if username is None:
            username = events[0].user
        print(username)
        print("\tAverage hours per month:", format(total_hours/len(by_month_hours), '.2f'))

        print("\tTotal hours by month:")
        for month in sorted(by_month_hours.keys()):
            print("\t", month, format(by_month_hours[month], '.2f'))

        print("\tMeetings per month:")
        for month in sorted(by_month_events.keys()):
            print("\t", month, by_month_events[month])

        print("\tAttendees per meeting per month:")
        for month in sorted(by_month_events.keys()):
            num_attendees = by_month_attendees[month]
            meetings = by_month_events[month]
            if num_attendees > 0:
                print("\t", month, format(1.0 * num_attendees/meetings, '.2f'))

        print("\n")

if __name__ == '__main__':
    file_dir = sys.argv[1]
    all_files = get_files(file_dir)
    all_events = []
    for file in all_files:
        events = read_file(os.path.join(file_dir, file))
        filtered_events = filter(events)
        analyze(filtered_events)
        all_events.extend(events)
    analyze(all_events, 'All')
