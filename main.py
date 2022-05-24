import win32com.client as client
from bs4 import BeautifulSoup as bs
import strings
import itertools


outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

def get_folder_by_name(folder_name, index):
    for folder in mapi.Folders.Item(index).Folders: 
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

    print_work_days(user, morning_shift_MSGs, morning_shift_LSWS, week_dates)
    print_work_days(user, evening_shift_MSGs, evening_shift_LSWS, week_dates)

def print_work_days(user, msgs_shift, lsws_shift, week_dates):

    days_i_work_MSGs = [i for i, s in enumerate(msgs_shift) if user in s] # Get the index where Eran shown in MSGs
    days_i_word_LSWS = [i for i, s in enumerate(lsws_shift) if user in s] # Get the index where Eran shown in LSWS

    for i, j in itertools.zip_longest(days_i_work_MSGs, days_i_word_LSWS):
        if i != None:
            print(week_dates[i - 1], strings.morning_monitor_job)
        
        if j != None:
            print(week_dates[j - 1], strings.morning_LS_WS_job)

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
    

    

def main():
    for i in range(1,4):
        folder =  get_folder_by_name("Work Schedule", i)

        if folder != None:
            latest_mail = list(folder.Items)[0]
            print('Subject: ', latest_mail.subject)
            print()
            parse_mail(latest_mail, 'Eran')
            return
        
    print('ERROR 404: Folder not found')



if __name__ == '__main__':
    main()