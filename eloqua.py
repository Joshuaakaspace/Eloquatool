import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('BluePurple')

EXCEL_FILE = 'Book1.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(13,2)), sg.InputText(key='Name')],
    [sg.Text('Eloqua', size=(13,2)), sg.InputText(key='Eloqua')],
    [sg.Text('Region', size=(13,2)), sg.Combo(['APAC', 'EMEA', 'AMERS'], key='Region')],
    [sg.Text('Date and Time', size=(13,2)), sg.InputText(key='Date and Time'),sg.CalendarButton('Date and Time')],
    [sg.Text('Emp ID', size=(13,2)), sg.InputText(key='Emp ID')],
  

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Eloqua Tracker Application', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
