import os
import win32com.client as win32
import datetime as dt
from views import dates_list, map
from convertor import dirlist_sorting

def recipients_generator(client_emails):
    rec_set = set()
    rec_list = []
    for client in client_emails:
        if client not in rec_set:
            rec_set.add(client)
            rec_list.append(client)
    return rec_list, rec_set

def create_mail(client_names, dates, recipient, attachments, carer, send=True):
    for client, rec, file in zip(client_names, recipient, attachments):
        client_str = "/".join(client)
        try:
            outlook = win32.Dispatch('outlook.application')
        except:
            return "Outlook doesn't exist, make sure it is downloaded correctly"
        mail = outlook.CreateItem(0)
        mail.To = rec
        mail.Subject = f'{carer} Services Invoice {dates[-1]}'
        mail.HtmlBody = f"Hello,<br><br>Here is the invoice for the week starting {dates[0]} for {client_str}<br><br>{carer}<br>0447 577 179"
        for f in file:
            mail.Attachments.Add(os.path.join(os.getcwd(),f'Invoices/{f}'))
        if send:
            mail.send()
        else:
            mail.save()
    return "Successfully created email drafts"

def attachments_generator(rec_list, rec_set, files, client_emails):
    clients, attachments_list = []
    attachment_set = set()
    attachments_list = []
    for k, v in zip(map["Emails"].keys(), map["Emails"].values()):
        if v in rec_set and k in map["Worked with"].keys():
            idx = rec_list.index(v)
            for file in files:
                if file[-15] == k[0] and file[-14] == k.split(' ')[-1][0]:
                    invoice = file
            if v == client_emails[idx] and v in attachment_set:
            # if len(attachments_list) > idx + 1:
                attachments_list[idx].append(invoice)
                clients[idx].append(k)
            else:
                attachment_set.add(v)
                attachments_list.append([invoice])
                clients.append([k])
    return attachments_list, clients, map["Carer"]

def main():
    fileDir = os.path.dirname(os.path.realpath('__file__'))
    files = []
    file_dupe_check = set()
    today = dt.datetime.now().date()
    for file in dirlist_sorting(fileDir):
        filetime = dt.datetime.fromtimestamp(
                os.path.getctime('Invoices/' + file))
        # add this optimised checking into dirlist_sorting?
        if len(file_dupe_check) == len(map["Name"]):
            break
        if filetime.date() == today and file[-15:-13] not in file_dupe_check:
            files.append(file)
            file_dupe_check.add(file[-15:-13])
    
    client_names = [key for key in map["Worked with"].keys()]
    client_emails = [map["Emails"][x] for x in client_names]
    dates = [dates_list[0], dates_list[-1]]
    recipients, rec_set = recipients_generator(client_emails)
    attachments, clients, carer = attachments_generator(recipients, rec_set, files, client_emails)
    message = create_mail(clients, dates, recipients, attachments, carer, send=False)
    return message