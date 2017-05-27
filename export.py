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

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar Meeting Room Shaming'

MAX_EVENTS = 1000

Event = namedtuple('Event', ['start', 'end', 'summary', 'creator', 'declined', 'start_timestamp', 'end_timestamp', 'attendees'], verbose=False)

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
        end = event['end'].get('dateTime', event['end'].get('date'))
        summary = event['summary']
        creator = event['creator']['email']

        start_timestamp = arrow.get(start).timestamp
        end_timestamp = arrow.get(end).timestamp

        declined = False
        attendees = []
        if 'attendees' in event:
            for attendee in event['attendees']:
                declined = attendee.get('resource', False) and \
                    attendee['responseStatus'] == 'declined' and \
                    attendee.get('self', True)

                if 'resource' not in attendee and \
                    attendee['responseStatus'] in ('tentative', 'accepted'):
                    attendees.append(attendee['email'])

        # if len(attendees) > 10:
        #     print(json.dumps(event, indent=2))
        #     exit()

        return Event(start, end, summary, creator, declined,
            start_timestamp, end_timestamp, attendees)
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

    print('Found rooms: ', ', '.join(cal['summary'] for cal in rooms))

    all_events = []
    for room in rooms:
        print('Getting events for', room['summary'])
        events_result = service.events().list(
            calendarId=room['id'], timeMin=now, maxResults=MAX_EVENTS, singleEvents=True,
            orderBy='startTime').execute()
        events = events_result.get('items', [])
        all_events.extend([parse_event(event) for event in events])

    print(json.dumps(all_events, indent=2))

    if outfile:
        with open(outfile, 'w') as f:
            writer = csv.writer(f)
            writer.writerow(('start', 'end', 'summary', 'creator', 'attendees'))
            writer.writerows([(e.start, e.end, e.summary, e.creator, '|'.join(e.attendees)) \
                for e in all_events if e is not None])

if __name__ == '__main__':
    main('out.csv')
