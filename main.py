from __future__ import print_function

from datetime import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from bs4 import BeautifulSoup as bs
import win32com.client as client
import strings
import itertools

######################################## Parse Mail ########################################
OUTLOOK = client.Dispatch("Outlook.Application")
MAPI = OUTLOOK.GetNamespace("MAPI")
working_days_info = {}

def get_folder_by_name(folder_name, index):
    for folder in MAPI.Folders.Item(index).Folders: 
        if folder.Name == folder_name:
            return folder
    
    return None

def parse_mail(mail, user):
    body = mail.HTMLBody
    soup = bs(body,'html.parser')
    mail_table = soup.find('table', {'class':'MsoNormalTable'})
    table_tds = [x.getText() for x in mail_table.find_all('td')]

    week_dates = table_tds[1:7]
    morning_shift_MSGs, morning_shift_LSWS, evening_shift_MSGs, evening_shift_LSWS = get_worker_lists(table_tds)

    create_work_days_dict(user, morning_shift_MSGs, morning_shift_LSWS, week_dates, 'morning')
    create_work_days_dict(user, evening_shift_MSGs, evening_shift_LSWS, week_dates, 'evening')

def create_work_days_dict(user, msgs_shift, lsws_shift, week_dates, shift):
    days_i_work_MSGs = [i for i, s in enumerate(msgs_shift) if user in s] # Get the index where Eran shown in MSGs
    days_i_word_LSWS = [i for i, s in enumerate(lsws_shift) if user in s] # Get the index where Eran shown in LSWS
    keys = ["Start_Date_Time", "End_Date_Time", "Role"]
    work_count = len(working_days_info)
    
    year = str(datetime.now().year)    

    for i, j in itertools.zip_longest(days_i_work_MSGs, days_i_word_LSWS):
        if i != None:
            date = year + '-' + week_dates[i-1]
            date = str(datetime.strptime(date,'%Y-%d-%b').date())
            
            if shift == 'morning':
                values = [date + 'T05:45:00+03:00', date + 'T15:00:00+03:00', 'Monitoring and MSGs']

            else:
                values = [date + 'T14:45:00+03:00', date + 'T23:00:00+03:00', 'Monitoring and MSGs']
            
            working_days_info[work_count] = {i:j for i,j in zip(keys, values)}
            work_count += 1
        
        
        if j != None:
            date = year + '-' + week_dates[j-1]
            date = str(datetime.strptime(date,'%Y-%d-%b').date())

            if shift == 'morning':
                values = [date + 'T05:45:00+03:00', date + 'T15:00:00+03:00', 'LS + WS']
            else:
                values = [date + 'T14:45:00+03:00', date + 'T23:00:00+03:00', 'LS + WS']

            working_days_info[work_count] = {i:j for i,j in zip(keys, values)}
            work_count += 1
            
def get_worker_lists(table_tds):
    morning_shift_monitor_index = table_tds.index(strings.morning_monitor_job)
    morning_shift_LSWS_index = table_tds.index(strings.morning_LS_WS_job)
    evening_shift_monitor_index = table_tds.index(strings.evening_monitor_job)
    evening_shift_LSWS_index = table_tds.index(strings.evening_LS_WS_job)

    morning_shift_MSGs = table_tds[morning_shift_monitor_index : morning_shift_LSWS_index]
    morning_shift_LS = table_tds[morning_shift_LSWS_index : morning_shift_LSWS_index + 7]
    evening_shift_MSGs = table_tds[evening_shift_monitor_index : evening_shift_LSWS_index]
    evening_shift_LS = table_tds[evening_shift_LSWS_index : evening_shift_LSWS_index + 7]

    return morning_shift_MSGs, morning_shift_LS, evening_shift_MSGs, evening_shift_LS
    
######################################## Calendar API ########################################
SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_event_info(appointment_info):
    event = {
        'summary':'Work !!',
        'location':'נמל אשדוד החדש',
        'description':f'Get ready to work! you are on the {appointment_info["Role"]}',
        'start':{'dateTime':f'{appointment_info["Start_Date_Time"]}', 'timeZone': 'Asia/Jerusalem'},
        'end':{'dateTime':f'{appointment_info["End_Date_Time"]}', 'timeZone':'Asia/Jerusalem'},
        'reminders':{'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 24*60}]}
        }

    return event

def get_credentials():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def create_event():
    creds = get_credentials()

    try:
        service = build('calendar', 'v3', credentials=creds)

        for i in working_days_info:
            event = get_event_info(working_days_info[i])
            print(event)
            event = service.events().insert(calendarId='primary', body=event).execute()
       
            print('Event created: %s \n' % (event.get('htmlLink')))

    except HttpError as error:
        print('An error occurred: %s' % error)
    
############################################ Main ############################################
def main():
    folder_index = 1

    while folder_index:
        folder =  get_folder_by_name("Work Schedule", folder_index)

        if folder != None:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            latest_mail = list(items)[0]
            
            print(f'Subject: {latest_mail.subject} \n')

            parse_mail(latest_mail, 'Eran')
            create_event()
            return

        folder_index += 1
        
    print('ERROR 404: Folder not found')



if __name__ == '__main__':
    main()