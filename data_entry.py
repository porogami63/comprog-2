from pathlib import Path
import PySimpleGUI as sg
import pandas as pd

#
sg.theme('Black')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'

if EXCEL_FILE.exists():
    df = pd.read_excel(EXCEL_FILE)
else:
    df = pd.DataFrame()

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Address', size=(15,1)), sg.InputText(key='Address')],
    [sg.Text('Year Level', size=(15,1)), sg.Combo(['1', '2', '3', '4', '5'], key='Year Level')],
    [sg.Text('College', size=(15,1)),
                            sg.Checkbox('CCS', key='CCS'),
                            sg.Checkbox('CEA', key='CEA'),
                            sg.Checkbox('COA', key='COA')],
    [sg.Text('Number of Units', size=(15,1)), sg.Spin([i for i in range(0,24)],
                                                       initial_value=0, key='Units')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('UNIVERSITY STUDENT DATA ENTRY FORM', layout)

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
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)  
        sg.popup('Data saved!')
        clear_input()
window.close()
