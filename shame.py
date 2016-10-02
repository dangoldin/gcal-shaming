from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime, json
import arrow

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'

MAX_EVENTS = 1000

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

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'

    calendars = service.calendarList().list(maxResults=200).execute()
    rooms = []
    for calendar in calendars['items']:
        if 'The' in calendar['summary']:
            print(calendar['summary'], calendar)
            rooms.append(calendar)

    people_meetings = {}

    for room in rooms:
        print('Getting events for', room['summary'])
        eventsResult = service.events().list(
            calendarId=room['id'], timeMin=now, maxResults=MAX_EVENTS, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        for event in events:
            if 'summary' in event:
                start = event['start'].get('dateTime', event['start'].get('date'))
                end  = event['end'].get('dateTime', event['end'].get('date'))
                summary = event['summary']
                creator = event['creator']['email']
                # print(json.dumps(event, indent=2))

                declined = False
                if 'attendees' in event:
                    for attendee in event['attendees']:
                        if (attendee.get('resource', False) and
                            attendee['responseStatus'] == 'declined' and
                            attendee.get('self', True)):
                                declined = True

                if declined:
                    continue

                # print("\t".join([start, end, room['summary'], creator, summary]))

                start_timestamp = arrow.get(start).timestamp
                end_timestamp = arrow.get(end).timestamp

                if creator not in people_meetings:
                    people_meetings[creator] = {}

                if (start_timestamp, end_timestamp) in people_meetings[creator]:
                    people_meetings[creator][(start_timestamp, end_timestamp)].append(summary + ' @ ' + room['summary'])
                    print('Found duplicate', creator, start, end, ':', ', '.join(people_meetings[creator][(start_timestamp, end_timestamp)]))
                else:
                    people_meetings[creator][(start_timestamp, end_timestamp)] = [summary + ' @ ' + room['summary']]

    # print(people_meetings)

if __name__ == '__main__':
    main()
