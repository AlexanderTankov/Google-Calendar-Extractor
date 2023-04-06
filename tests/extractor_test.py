import unittest
from unittest.mock import MagicMock
import pytz
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from download_calendar_events import download_events


class TestDownloadEvents(unittest.TestCase):

    def setUp(self):
        self.service_mock = MagicMock(spec=build('calendar', 'v3'))
        self.creds_mock = MagicMock(spec=Credentials)
        self.start_date = datetime(2023, 3, 1)
        self.end_date = datetime(2023, 3, 31)

    def test_download_events_no_events(self):
        events_result_mock = MagicMock()
        events_result_mock.get.return_value = []
        self.service_mock.events().list().execute.return_value = events_result_mock

        with self.assertLogs() as logs:
            download_events(self.start_date, self.end_date)

        self.assertEqual(logs.output, ['INFO:root:No events found.'])

    def test_download_events_one_event(self):
        event = {
            'summary': 'Test event',
            'start': {
                'dateTime': '2023-03-02T10:00:00-08:00',
                'timeZone': 'America/Los_Angeles'
            },
            'end': {
                'dateTime': '2023-03-02T11:00:00-08:00',
                'timeZone': 'America/Los_Angeles'
            }
        }
        events_result_mock = MagicMock()
        events_result_mock.get.return_value = [event]
        self.service_mock.events().list().execute.return_value = events_result_mock

        with self.assertLogs() as logs:
            download_events(self.start_date, self.end_date)

        self.assertEqual(logs.output, [
            'INFO:root:Events from 03/01/2023 to 03/31/2023:',
            'INFO:root:Test event at 10:00 AM'
        ])

    def test_download_events_multiple_events(self):
        event1 = {
            'summary': 'Test event 1',
            'start': {
                'dateTime': '2023-03-02T10:00:00-08:00',
                'timeZone': 'America/Los_Angeles'
            },
            'end': {
                'dateTime': '2023-03-02T11:00:00-08:00',
                'timeZone': 'America/Los_Angeles'
            }
        }
        event2 = {
            'summary': 'Test event 2',
            'start': {
                'dateTime': '2023-03-05T15:30:00-08:00',
                'timeZone': 'America/Los_Angeles'
            },
            'end': {
                'dateTime': '2023-03-05T17:00:00-08:00',
                'timeZone': 'America/Los_Angeles'
            }
        }
        events_result_mock = MagicMock()
        events_result_mock.get.return_value = [event1, event2]
        self.service_mock.events().list().execute.return_value = events_result_mock

        with self.assertLogs() as logs:
            download_events(self.start_date, self.end_date)

        self.assertEqual(logs.output, [
            'INFO:root:Events from 03/01/2023 to 03/31/2023:',
            'INFO:root:Test event 1 at 10:00 AM',
            'INFO:root:Test event 2 at 03:30 PM'])