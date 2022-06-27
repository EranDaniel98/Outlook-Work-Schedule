import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Configs.config_handler import config_handler
import strings


######################################## Calendar API ########################################

class Calendar_API:
    def __init__(self, working_days_info):
        self.user_config_file = config_handler('user_config.yaml')
        self.software_config_file = config_handler('config.yaml')

        self.SCOPES = self.software_config_file.get_requested_param('SCOPES')
        self.working_days_info =  working_days_info

    def create_event_info(self, shit_info):
        event =  self.user_config_file.get_requested_param('event')
        if event == None or len(event) < 3: return None # write to log 
        
        event['description'] = strings.event_description + shit_info["Role"]
        event['start'] = {'dateTime': f'{shit_info["Start_Date_Time"]}', 'timeZone': 'Asia/Jerusalem'} 
        event['end'] = {'dateTime': f'{shit_info["End_Date_Time"]}', 'timeZone':'Asia/Jerusalem'} 

        return event

    def get_OAuth_credentials(self):
        token_path = self.software_config_file.get_requested_param('token_path')
        cred_path = self.software_config_file.get_requested_param('cred_path')
        creds = None

        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, self.SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(cred_path, self.SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open(token_path, 'w') as token:
                token.write(creds.to_json())
        
        return creds

    def create_event(self):
        print(self.working_days_info)
        creds = self.get_OAuth_credentials()

        try:
            service = build('calendar', 'v3', credentials=creds)

            for i in self.working_days_info:
                event = self.create_event_info(self.working_days_info[i])
                if event == None: continue
                
                print(event)
                event = service.events().insert(calendarId='primary', body=event).execute()
        
                print('\nEvent created: %s \n' % (event.get('htmlLink'))) # add to log

        except HttpError as error:
            print('An error occurred: %s' % error)