import PyQt5 as pq
import sys
from PyQt5 import QtWidgets, uic, QtGui, QtCore
import datetime
import os.path
import ctypes

map = {
    "Worked with": {},
    "person": 0,
    "Name": ["Elwyn Bourke", "Jonathan McNeill", "Zenith Burke"],
    "KMs": ["0", "0", "0"],
    "Monday": {
        "PH": [False, False, False],
        "Hours": ["1", "1", "0"]},
    "Tuesday": {
        "PH": [False, False, False],
        "Hours": ["0", "1", "0"]},
    "Wednesday": {
        "PH": [False, False, False],
        "Hours": ["1", "1", "0"]},
    "Thursday": {
        "PH": [False, False, False],
        "Hours": ["1", "1", "0"]},
    "Friday": {
        "PH": [False, False, False],
        "Hours": ["0", "1", "0"]},
    "Saturday": {
        "PH": [False, False, False],
        "Hours": ["0", "0", "0"]},
    "Sunday": {
        "PH": [False, False, False],
        "Hours": ["0", "0", "0"]}
}

# GUI functionality
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# dynamically grabbing date of x days ago
def new_date(today, number):
    dt = datetime.timedelta(days = number)
    return today - dt

# GUI functionality

# function for confirming hours
def confirmation_setup():
    c_dlg.show()
    string = f"You are about to submit {len(map['Worked with'])} invoices with the following information:\n"
    for name in map["Worked with"]:
        string += f'\n{map["Worked with"][name][0]} Hour/s for {name} ({", ".join(map["Worked with"][name][1])})' 
    c_dlg.text_label.setText(string)

# main_view GUI functions
# function for updating text
def text_updating(labels, idxs=None, col=None):
    global dates_list
    global days_list
    # check if updating 'days' or default value in 'hours' input
    if idxs:
        for (label, day, idx) in zip(labels, days_list, idxs):
            label.setText(f'{day} ({dates_list[idx][:5]})') 
    else:
        for (label, day) in zip(labels, days_list):
            label.setText(map[day][col][map["person"]])

# function for updating isChecked of passed checkboxes
def checkbox_values(labels, col):
    global days_list
    for (label, day) in zip(labels, days_list):
        label.setChecked(map[day][col][map["person"]]) 

# master function for setting values when called
def setting_view():
    # heading label
    call.heading.setText(f'{dates_list[0]} -> {dates_list[-1]} - {map["Name"][map["person"]]}')

    # days col 
    text_updating(labels=[call.label_sun, call.label_mon, call.label_tues, call.label_wed, call.label_thurs, call.label_fri, call.label_sat], idxs=list(range(0,len(days_list)+1)))
    
    # Hours input
    text_updating(labels=[call.input_w_su, call.input_w_m, call.input_w_tu, call.input_w_w, call.input_w_th, call.input_w_f, call.input_w_sa], col="Hours")
    
    # PH checkboxes
    checkbox_values([call.cb_ph_su, call.cb_ph_m, call.cb_ph_tu, call.cb_ph_w, call.cb_ph_th, call.cb_ph_f, call.cb_ph_sa], "PH")
    
    call.submitButton.setText("Next" if map["person"] < len(map["Name"]) - 1 else "Submit")
    call.prevButton.setVisible(map["person"] != 0)
    call.input_kms.setMaximum(999)

# function for updating map
def update_map(day, variables):
    map[day]["PH"][map["person"]] = variables[0].isChecked()
    map[day]["Hours"][map["person"]] = variables[1].text()
    if float(variables[1].text()) > 0:
        map["Worked with"][map["Name"][map["person"]]][0] += float(variables[1].text())
        map["Worked with"][map["Name"][map["person"]]][1].append(day[:2])

# function for checking validity of inputted values
def input_check(hours_labels, PH_labels):
    count = 0
    for l1, l2 in zip(hours_labels, PH_labels):
        try:
            float(l1.text())
        except ValueError:
            return "Please input either integer (e.g., 1) or float (e.g., 1.5) values into the hours column"
        count += 1 if l2.isChecked() else 0
    return 'There can only be a maximum of 3 public holiday shifts per client per week' if count > 3 else ''

# Functionality for prev button
def Previous():
    map["person"] -= 1
    if map["Name"][map["person"]] in map["Worked with"].keys():
        del map["Worked with"][map["Name"][map["person"]]]
    setting_view()

# functionality of skip button
def Next_client():
    map["person"] += 1
    if map["person"] < len(map["Name"]):
        setting_view()
        return
    confirmation_setup()

# functionality of submit button
def Submit(func=None):
    message = input_check([call.input_w_m, call.input_w_tu, call.input_w_w, call.input_w_th, call.input_w_f, call.input_w_sa, call.input_w_su],
                    [call.cb_ph_m, call.cb_ph_tu, call.cb_ph_w, call.cb_ph_th, call.cb_ph_f, call.cb_ph_sa, call.cb_ph_su])

    if message:
        ctypes.windll.user32.MessageBoxW(0, message, 1)
        return
    
    if func == 'skip':
        Next_client()
        return
    
    if map["person"] < len(map["Name"]):
    # update map with data, for each day 
        map["Worked with"][map["Name"][map["person"]]] = [0, []]
        update_map(days_list[0], [call.cb_ph_su, call.input_w_su])
        update_map(days_list[1], [call.cb_ph_m, call.input_w_m])
        update_map(days_list[2], [call.cb_ph_tu, call.input_w_tu])
        update_map(days_list[3], [call.cb_ph_w, call.input_w_w])
        update_map(days_list[4], [call.cb_ph_th, call.input_w_th])
        update_map(days_list[5], [call.cb_ph_f, call.input_w_f])
        update_map(days_list[6], [call.cb_ph_sa, call.input_w_sa])

        if map["Worked with"][map["Name"][map["person"]]][0] == 0:
            del map["Worked with"][map["Name"][map["person"]]]
        # Update map with KMs
        map["KMs"][map["person"]] = call.input_kms.text()

        if map["person"] < len(map["Name"]) - 1:
            Next_client()
            return
    confirmation_setup()

# function for yes button on confirmation popup
def Dlg_Submit():
    global map_update
    map["person"] = 0
    c_dlg.close()
    call.close()
    map_update = True

if __name__ == "views":
    date_sat = datetime.datetime.today()
    date_idx = date_sat.weekday()
    if date_idx != 5:
        date_sat = new_date(date_sat, date_idx - 5) if date_idx - 5 > 0 else new_date(date_sat, date_idx + 2)
    
    # getting list of all dates in last week
    dates_list = [new_date(date_sat, 6-x).strftime("%d/%m/%Y") for x in range(7)]
    # map_update = False
    days_list = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    # Loading windows into memory from main_view.ui and confirmation_popup.ui
    app=QtWidgets.QApplication([])
    c_dlg=pq.uic.loadUi(resource_path("Views/confirmation_popup.ui"))
    call=pq.uic.loadUi(resource_path("Views/main_view.ui"))

    setting_view()
    call.submitButton.clicked.connect(Submit)
    call.prevButton.clicked.connect(Previous)
    call.skipButton.clicked.connect(lambda: Submit('skip'))

    c_dlg.noButton.clicked.connect(c_dlg.close)
    c_dlg.yesButton.clicked.connect(Dlg_Submit)

    # Opening window from memory
    call.show()
    app.exec_()
