import win32com.client as client
from bs4 import BeautifulSoup as bs


outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

def get_folder_by_name(folder_name):
    for folder in mapi.Folders.Item(1).Folders: 
        if folder.Name == folder_name:
            return folder
    
    return None

def parse_mail(mail, user):
    body = mail.HTMLBody
    soup = bs(body,'html.parser')
    mail_table = soup.find('table', {'class':'MsoNormalTable'})
    table_tds = [x.getText() for x in mail_table.find_all('td')]

    week_dates = table_tds[1:7]
    morning_shift_index = table_tds.index('מנטורינג + מסרים 05:45-15:00') + 1
    morning_shift_index2 = table_tds.index('LS + WS 05:45-15:00') + 1
    evening_shift_index = table_tds.index('מנטורינג + מסרים 14:45-23:00') + 1
    training_shift_index = table_tds.index('08:00-17:00')

    morning_shift_workers = table_tds[morning_shift_index:morning_shift_index2 + 6]
    evening_shift_workers = table_tds[evening_shift_index:training_shift_index]
    print(week_dates)
    print(len(morning_shift_workers), morning_shift_workers)
    #print()
    #print(len(evening_shift_workers), evening_shift_workers)


    days_i_work = [i for i, s in enumerate(morning_shift_workers) if user in s] # Get the index where Eran shown
    print(days_i_work)
    for i in days_i_work:
        if i >= 0 and i < 6:
            print(week_dates[i], '05:45-15:00 Monitoring and MSGs')

        if i > 6 and i < len(morning_shift_workers): 
            print(week_dates[len(morning_shift_workers) - i + 1], '05:45-15:00 LS + WS')
    
    days_i_work = [i for i, s in enumerate(evening_shift_workers) if user in s] # Get the index where Eran shown
    for i in  days_i_work:
        if i >= 0 and i < 6: 
            print(week_dates[i], '14:45-23:00 Monitoring and MSGs')
        
        if i > 6 and i < len(evening_shift_workers): 
            print(week_dates[i-7], '14:45-23:00 LS + WS')



def main():
    folder =  get_folder_by_name("Work Schedule")

    if folder == None:
        print('ERROR 404: Folder not found')
        return

    #print(i.SentOn.strftime("%d-%m-%y")) - Date
    latest_mail = list(folder.Items)[1]
    #for i in folder.Items:
    print(latest_mail.subject)
    parse_mail(latest_mail, 'Eran')



if __name__ == '__main__':
    main()