import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from mail_parser import mail_parser


######################################## Calendar API ########################################

class Calendar_API:
    def __init__(self, working_days_info):
        self.SCOPES = ['https://www.googleapis.com/auth/calendar']
        self.working_days_info =  working_days_info

    def create_event_info(self, appointment_info):
        event = {
            'summary':'Work !!',
            'location':'נמל אשדוד החדש',
            'description':f'Get ready to work! you are on the {appointment_info["Role"]}',
            'start':{'dateTime':f'{appointment_info["Start_Date_Time"]}', 'timeZone': 'Asia/Jerusalem'},
            'end':{'dateTime':f'{appointment_info["End_Date_Time"]}', 'timeZone':'Asia/Jerusalem'},
            'reminders':{'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 24*60}]}
            }

        return event

    def get_credentials(self):
        creds = None

        if os.path.exists('GoogleAPI/token.json'):
            creds = Credentials.from_authorized_user_file('GoogleAPI/token.json', self.SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'GoogleAPI/credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open('GoogleAPI/token.json', 'w') as token:
                token.write(creds.to_json())
        
        return creds

    def create_event(self):
        print(self.working_days_info)
        creds = self.get_credentials()

        try:
            service = build('calendar', 'v3', credentials=creds)

            for i in self.working_days_info:
                event = self.create_event_info(self.working_days_info[i])
                print(event)
                event = service.events().insert(calendarId='primary', body=event).execute()
        
                print('\nEvent created: %s \n' % (event.get('htmlLink')))

        except HttpError as error:
            print('An error occurred: %s' % error)