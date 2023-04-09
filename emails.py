import os
import json
import sys
import win32com.client as win32
from views import dates_list, resource_path, map

# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller """
#     try:
#         # PyInstaller creates a temp folder and stores path in _MEIPASS
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)

def recipients_generator(client_emails):
    set = set()
    rec_list = []
    for client in client_emails:
        if client not in set:
            set.add(client)
            rec_list.append(client)
    return rec_list, set

def create_mail(client_names, dates, recipient, attachments, send=True):

    for client, rec, file in zip(client_names, recipient, attachments):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = rec
        mail.Subject = f'Kita Liosatos Services Invoice {dates[-1]}'
        mail.HtmlBody = f"Hello,<br><br>Here is the invoice for the week starting {dates[0]} for {client}<br><br>Kita Liosatos<br>0447 577 179"
        for f in file:
            mail.Attachments.Add(os.path.join(os.getcwd(),f'Invoices/{f}'))
        if send:
            mail.send()
        else:
            mail.save()

def attachments_generator(rec_list, rec_set):
    attachments_list = []
    for k, v in map["Emails"]:
        if v in rec_set and k in map["Worked with"].keys():
            idx = rec_list.index(v)
            if len(attachments_list) > idx + 1:
                attachments_list[idx].append(v)
            else:
                attachments_list.append([v])
    return attachments_list
client_names = [key for key in map["Worked with"].keys()]
client_emails = [map[x] for x in client_names]
dates = [dates_list[0], dates_list[-1]]
recipients, rec_set = recipients_generator(client_emails)
attachments = attachments_generator(recipients, rec_set)

# attachments = [['Invoice_1EB_08042023.pdf'], ['Invoice_2JM_08042023.pdf']]
create_mail(client_names, dates, recipients, attachments, send=False)



# def main(client_dict):
#     global fileDir, xcl_file, ar, wb, ws, message
#     # Excel file path information
#     fileDir = os.path.dirname(os.path.realpath('__file__'))
#     xcl_file = 'Invoice-Template.xlsx'

#     # Make invoice folder
#     if not os.path.exists(f'{fileDir}\Invoices'):
#         os.makedirs(f'{fileDir}\Invoices')

#     # setting constant alignment variable
#     ar = Alignment(horizontal='right')

#     # importing worksheet
#     try:
#         wb = openpyxl.load_workbook(xcl_file)
#         ws = wb.worksheets[0]
#     except:
#         message = f"Sorry, we couldn't find '{xcl_file}'. Is it possible it was moved, renamed or deleted?"
        
#     if not message:
#         message = xc2pdf([x for x in client_dict["Worked with"]], client_dict)
#     return message