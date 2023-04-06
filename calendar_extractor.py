import pytz

from datetime import datetime
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build


class CalendarExtractor:
    ''' CalendarExtractor will be responsible for getting the information from the google calendar of
    Silvy Spa through the Google Clould API, structure the events and provide them in a file'''

    def __init__(self, start_date, end_date, creds) -> None:
        self.start_date = start_date
        self.end_date = end_date
        self.creds = creds

    def download_events(self):
        ''' Download the events from the google calendar and return them in a list. '''
        # Set up the Calendar API client
        service = build('calendar', 'v3', credentials=self.creds)

        # Set the time zone for the start and end dates
        tz = pytz.timezone('Europe/Sofia')

        # Convert the start and end dates to UTC
        start_date_utc = tz.localize(self.start_date).astimezone(pytz.utc).strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        end_date_utc = tz.localize(self.end_date).astimezone(pytz.utc).strftime('%Y-%m-%dT%H:%M:%S.%fZ')

        events = []
        try:
            # Call the Calendar API to get events
            events_result = service.events().list(calendarId='primary', timeMin=start_date_utc, timeMax=end_date_utc, maxResults=1000, singleEvents=True, orderBy='startTime').execute()
            events = events_result.get('items', [])
        except HttpError as error:
            print('An error occurred: {}'.format(error))
        return events

    def store_events_in_workbook(self, events, workbook):
        ''' Store the events to the workbook and save it.'''
        num_stored_events = 0
        if not events:
            print('No events to store.')
        else:
            print('Events from {} to {}:'.format(self.start_date.strftime('%m/%d/%Y'), self.end_date.strftime('%m/%d/%Y')))
            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                start_time = datetime.fromisoformat(start).strftime('%I:%M %p')
                print('Stored {} at {}'.format(event['summary'], start_time))
                workbook.add_event(event)
                num_stored_events += 1
        workbook.save()
        return num_stored_events
