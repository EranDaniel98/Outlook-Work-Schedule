from datetime import datetime
from bs4 import BeautifulSoup as bs
import win32com.client as client
import itertools
from Configs.config_handler import config_handler
import strings

OUTLOOK = client.Dispatch("Outlook.Application")
MAPI = OUTLOOK.GetNamespace("MAPI")

######################################## Parse Mail ########################################
class mail_parser:

    def __init__(self):
        self.working_days_info = {}
        #self.config_file_data.__init__('user_config.yaml')
        self.config_file_data = config_handler('user_config.yaml')

    def get_folder_by_name(self):
        folder_name = self.config_file_data.get_folder_name()
        if folder_name == None: return None

        folder_index = 1
        while(folder_index):
            for folder in MAPI.Folders.Item(folder_index).Folders: 
                if folder.Name == folder_name:
                    return folder

            folder_index += 1
        
        # Write to logs
        return None

    def parse_mail(self, mail):
        body = mail.HTMLBody
        soup = bs(body,'html.parser')
        mail_table = soup.find('table', {'class':'MsoNormalTable'})
        table_tds = [x.getText() for x in mail_table.find_all('td')]

        week_dates = table_tds[1:7]
        morning_shift_MSGs, morning_shift_LSWS, evening_shift_MSGs, evening_shift_LSWS = self.get_worker_lists(table_tds)

        self.create_work_days_dict(morning_shift_MSGs, morning_shift_LSWS, week_dates, 'morning')
        self.create_work_days_dict(evening_shift_MSGs, evening_shift_LSWS, week_dates, 'evening')

    def create_work_days_dict(self, msgs_shift, lsws_shift, week_dates, shift):
        user_name = self.config_file_data.get_user_name()

        keys = ["Start_Date_Time", "End_Date_Time", "Role"]
        work_count = len(self.working_days_info)
        
        days_i_work_MSGs = [i for i, s in enumerate(msgs_shift) if user_name in s] # Get the index where Eran shown in MSGs
        days_i_word_LSWS = [i for i, s in enumerate(lsws_shift) if user_name in s] # Get the index where Eran shown in LSWS
        
        year = str(datetime.now().year)    

        for i, j in itertools.zip_longest(days_i_work_MSGs, days_i_word_LSWS):
            if i != None:
                date = year + '-' + week_dates[i-1]
                date = str(datetime.strptime(date,'%Y-%d-%b').date())
                
                if shift == 'morning':
                    values = [date + 'T05:45:00+03:00', date + 'T15:00:00+03:00', strings.monitoring]

                else:
                    values = [date + 'T14:45:00+03:00', date + 'T23:00:00+03:00', strings.monitoring]
                
                self.working_days_info[work_count] = {i:j for i,j in zip(keys, values)}
                work_count += 1
            
            
            if j != None:
                date = year + '-' + week_dates[j-1]
                date = str(datetime.strptime(date,'%Y-%d-%b').date())

                if shift == 'morning':
                    values = [date + 'T05:45:00+03:00', date + 'T15:00:00+03:00', strings.LSWS]
                else:
                    values = [date + 'T14:45:00+03:00', date + 'T23:00:00+03:00', strings.LSWS]

                self.working_days_info[work_count] = {i:j for i,j in zip(keys, values)}
                work_count += 1
                
    def get_worker_lists(self, table_tds):
        morning_shift_monitor_index = table_tds.index(strings.morning_monitor_job)
        morning_shift_LSWS_index = table_tds.index(strings.morning_LS_WS_job)
        morning_shift_MSGs = table_tds[morning_shift_monitor_index : morning_shift_LSWS_index]
        morning_shift_LS = table_tds[morning_shift_LSWS_index : morning_shift_LSWS_index + 7]

        evening_shift_monitor_index = table_tds.index(strings.evening_monitor_job)
        evening_shift_LSWS_index = table_tds.index(strings.evening_LS_WS_job)
        evening_shift_MSGs = table_tds[evening_shift_monitor_index : evening_shift_LSWS_index]
        evening_shift_LS = table_tds[evening_shift_LSWS_index : evening_shift_LSWS_index + 7]

        return morning_shift_MSGs, morning_shift_LS, evening_shift_MSGs, evening_shift_LS
        