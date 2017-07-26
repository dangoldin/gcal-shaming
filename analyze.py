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
    events = []
    with open(file, 'r') as f:
        c = csv.reader(f)
        c.next() # Skip header
        for row in c:
            e = Event(user, *row)
            events.append(e)
    return events

def filter(events):
    # Remove events > 24 hours and no attendees
    filtered_events = [e for e in events if float(e.duration_hours) < 24 and len(e.attendees) > 0]
    return filtered_events

def analyze(events):
    total_hours = 0.0
    by_month = defaultdict(float)
    for e in events:
        month = e.start[:7]
        by_month[month] += float(e.duration_hours)
        total_hours += float(e.duration_hours)
    if len(events):
        print(events[0].user)
        print("\t", total_hours/len(by_month))
        # print("\t", by_month)
        print("\n")


if __name__ == '__main__':
    file_dir = sys.argv[1]
    all_files = get_files(file_dir)
    for file in all_files:
        events = read_file(os.path.join(file_dir, file))
        filtered_events = filter(events)
        analyze(filtered_events)
