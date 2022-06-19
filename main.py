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
    
    # Get last mail from
    work_sched_folder =  mail_handler.get_work_sched_folder()
    latest_mail = get_last_mail(work_sched_folder)

    # Parse the mail    
    print(f'Subject: {latest_mail.subject} \n')
    mail_handler.parse_mail(latest_mail)
    
    # Create event/s
    even_creator = Calendar_API(mail_handler.working_days_info)
    even_creator.create_event()



if __name__ == '__main__':
    main()