import os.path
import win32com.client
from views import dates_list, days_list, resource_path

pay_information_dict = {
    "Weekday Support": ["15-045-0128-1-3", 45.0],
    "Saturday Support": ["01-013-0107-1-1", 70.0],
    "Sunday Support": ["01-014-0107-1-1", 75.0],
    "Public Holiday Support": ["01-012-0107-1-1", 90.0]
}

def Pdf_check(dirlist):
    try:
        file = dirlist[0].split('_')
    except:
        return False
    
    if (len(file) != 3) or (file[2][-3:] != 'pdf') or (file[0] != 'Invoice') or (len(file[2]) != 12):
        dirlist.pop(-1)
        file = Pdf_check(dirlist)
    if file:
        try:
            int(file[1][0:-2]) 
        except:
            dirlist.pop(0)
            file = Pdf_check(dirlist)
    return file

def dirlist_sorting(fileDir):
    dirlist = filter( lambda x: os.path.isfile(os.path.join(f'{fileDir}\Invoices', x)),
                                os.listdir(f'{fileDir}\Invoices') )
        # Sort list of files based on last modification time in ascending order
    dirlist = sorted( dirlist,
                            key = lambda x: os.path.getmtime(os.path.join(f'{fileDir}\Invoices', x))
                            )
    dirlist.reverse()
    return dirlist

# function to edit excel file
def Excel_edit(client_dict):    
    dirlist = dirlist_sorting(fileDir)
    
    if len(dirlist) > 0:
        file = Pdf_check(dirlist)
        if file:
            invoice_number = int(file[1][0:-2]) + 1
        else:
            invoice_number = 1
    else:
        invoice_number = 1

    # editing cells
    ws.Range('C9:C10').Value = [[dates_list[-1]], ["0"*(4-len(str(invoice_number))) + str(invoice_number)]]
    value = client_dict["Name"][client_dict["person"]]
    ws.Cells(12,3).Value = value
    cell_number = 16

    for x in days_list:
        # identifying if worked with this client on a particular day
        if float(client_dict[x]["Hours"][client_dict["person"]]) > 0:

            # identifying the type of shift (PH/weekend/weekday)
            if client_dict[x]["PH"][client_dict["person"]]:
                key = "Public Holiday Support"
            elif x == days_list[0]:
                key = "Sunday Support"
            elif x == days_list[-1]:
                key = "Saturday Support"
            else:
                key = "Weekday Support"
            
            value_list = [dates_list[days_list.index(x)], key, pay_information_dict[key][0], client_dict[x]["Hours"][client_dict["person"]], pay_information_dict[key][1]]
            # inputting correct information
            ws.Range(f'A{str(cell_number)}:E{cell_number}').Value = value_list
            cell_number += 1

    # blanking out rest of price table
    if ws.Cells(cell_number,1).Value: 
        ws.Range(f'A{str(cell_number)}:E25').ClearContents()

    # populating KMs cell
    ws.Cells(4,4).Value = client_dict["KMs"][client_dict["person"]] if int(client_dict["KMs"][client_dict["person"]]) > 0 else None
    return invoice_number, xcl_file

# master function, linking other functions together
def xc2pdf(clients, client_dict):
    for client in clients:
        client_dict["person"] = client_dict["Name"].index(client)
        if client in client_dict["Worked with"].keys():

            invoice_number, xcl_file = Excel_edit(client_dict)
            try:
                PATH_TO_PDF = f'{fileDir}\Invoices\Invoice_{str(invoice_number)}{client_dict["Name"][client_dict["person"]].split(" ")[0][0]}{client_dict["Name"][client_dict["person"]].split(" ")[1][0]}_{dates_list[-1].replace("/","")}.pdf'
                wb.SaveAs(f"{fileDir}\{xcl_file}")
                ws.ExportAsFixedFormat(0, PATH_TO_PDF)
            except:
                return f"Sorry, we couldn't find the output file ({fileDir}\Invoices\). Is it possible this folder was moved, renamed or deleted?"
    return ""

def main(client_dict):
    global fileDir, xcl_file, ws, wb
    # Excel file path information
    fileDir = os.path.dirname(os.path.realpath('__file__'))
    xcl_file = 'Invoice-Template.xlsx'
    WB_PATH = f'{fileDir}\{xcl_file}'
    if not os.path.exists(WB_PATH):
        print('no')
    else:
        print('yes')

    # Make invoice folder
    if not os.path.exists(f'{fileDir}\Invoices'):
        os.makedirs(f'{fileDir}\Invoices')
    
    # importing worksheet
    try:
        excelApp = win32com.client.Dispatch("Excel.Application")
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        wb = excelApp.Workbooks.Open(WB_PATH)
        ws = wb.Worksheets('Master')
    except:
        return f"Sorry, we couldn't find '{xcl_file}'. Is it possible it was moved, renamed or deleted?"
    
    message = xc2pdf([x for x in client_dict["Worked with"]], client_dict)
    wb.Close(True)
    excelApp.Quit()
    return message