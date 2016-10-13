from __future__ import print_function
import httplib2
import os
import sys
import csv

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime, json
import arrow
from rangeset import RangeSet
from collections import namedtuple

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar Meeting Room Shaming'

MAX_EVENTS = 1000

Event = namedtuple('Event', ['start', 'end', 'summary', 'creator', 'declined', 'start_timestamp', 'end_timestamp'], verbose=False)

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def parse_event(event):
    if 'summary' in event:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end  = event['end'].get('dateTime', event['end'].get('date'))
        summary = event['summary']
        creator = event['creator']['email']

        start_timestamp = arrow.get(start).timestamp
        end_timestamp = arrow.get(end).timestamp

        declined = False
        if 'attendees' in event:
            for attendee in event['attendees']:
                if (attendee.get('resource', False) and
                    attendee['responseStatus'] == 'declined' and
                    attendee.get('self', True)):
                        declined = True

        return Event(start, end, summary, creator, declined, start_timestamp, end_timestamp)
    return None

def main(outfile):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'

    # Make sure to get everyone
    calendars = service.calendarList().list(maxResults=200).execute()
    rooms = []
    for calendar in calendars['items']:
        # All our rooms start with 'The' so this needs to be cleaned up
        if 'The' in calendar['summary']:
            print(calendar['summary'], calendar)
            rooms.append(calendar)

    people_ranges = {}

    all_events = []
    for room in rooms:
        print('Getting events for', room['summary'])
        eventsResult = service.events().list(
            calendarId=room['id'], timeMin=now, maxResults=MAX_EVENTS, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        for event in events:
            details = parse_event(event)

            if not details:
                continue

            this_range = RangeSet(details.start_timestamp, details.end_timestamp)

            has_overlap = False
            if details.creator not in people_ranges:
                people_ranges[details.creator] = this_range
            else:
                if len(this_range & people_ranges[details.creator]):
                    print('Overlap meeting for',details)
                    has_overlap = True
                people_ranges[details.creator] |= this_range
            all_events.append((details, has_overlap))

    if outfile:
        with open(outfile, 'w') as f:
            w = csv.writer(f)
            w.writerow(('start', 'end', 'summary', 'creator', 'declined', 'start_timestamp', 'end_timestamp', 'has_overlap'))
            w.writerows([(e.start, e.end, e.summary, e.creator, e.declined, e.start_timestamp, e.end_timestamp, has_overlap) for e, has_overlap in all_events])

if __name__ == '__main__':
    outfile = None
    if len(sys.argv) > 1:
        outfile = sys.argv[1]

    main(outfile)
