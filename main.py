############################################ Main ############################################
from mail_parser import mail_parser
from GoogleAPI.API_handler import Calendar_API


def get_last_mail(folder):
    if folder != None:
        items = folder.Items
        items.sort("[ReceivedTime]", True)
        latest_mail = list(items)[0]
        return latest_mail
    
    return None

def main():
    mail_handler = mail_parser()

    folder =  mail_handler.get_folder_by_name()
    latest_mail = get_last_mail(folder)
        
    print(f'Subject: {latest_mail.subject} \n')
    mail_handler.parse_mail(latest_mail)
    
    even_creator = Calendar_API(mail_handler.working_days_info)
    even_creator.create_event()



if __name__ == '__main__':
    main()