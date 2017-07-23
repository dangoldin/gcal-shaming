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
from dateutil import tz
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
APPLICATION_NAME = 'Google Calendar Fun'

MAX_EVENTS = 1000
DAYS_BACK = 7 * 24 # 24 weeks
SECONDS_IN_DAY = 60 * 60 * 24

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
        summary = event['summary'].encode('utf-8')
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

        return Event(start, end, summary, creator, declined,
            start_timestamp, end_timestamp, attendees)
    return None

def main(outfile):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    all_events = []

    start_month = '2013-10-01'
    curr_datetime = arrow.get(start_month)

    while curr_datetime < arrow.now():
        end_datetime = curr_datetime.shift(months=1)
        start_ts = curr_datetime.format('YYYY-MM-DDTHH:mm:ss') + 'Z'
        end_ts = end_datetime.format('YYYY-MM-DDTHH:mm:ss') + 'Z'

        curr_datetime = end_datetime

        print('Getting events', start_ts, 'to', end_ts)
        try:
            # Docs: https://developers.google.com/google-apps/calendar/v3/reference/events/list
            events_result = service.events().list(
                calendarId='dgoldin@triplelift.com',
                timeMin=start_ts,
                timeMax=end_ts,
                maxResults=MAX_EVENTS, singleEvents=True,
                orderBy='startTime').execute()
            events = [parse_event(e) for e in events_result.get('items', []) if parse_event(e) is not None]
            all_events.extend(events)
        except Exception, e:
            print('\tFailed to get events', e)

    with open(outfile, 'w') as f:
        writer = csv.writer(f)
        writer.writerow(('start', 'end', 'duration_hours', 'summary', 'creator', 'attendees'))
        writer.writerows([(e.start, e.end, \
            (arrow.get(e.end) - arrow.get(e.start)).total_seconds()/(60.0 * 60.0), \
            e.summary, e.creator, '|'.join(e.attendees)) \
              for e in all_events if e is not None])

if __name__ == '__main__':
    out = 'all-meetings.csv'
    main(out)
