"""
use pysimplegui as an alternative to tkinter askopenfilename and askdirectory
USE THIS VERSION - V2 shortens one shot window to a few lines
"""
# taken from 'single shot' here: https://www.pysimplegui.org/en/latest/cookbook/

import os
import PySimpleGUI as sg


layout = [
    [sg.T("Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'")],
    [sg.Text("Choose a file: "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(initial_folder= os.path.expanduser("~/Downloads/"))
    ],
    [sg.Button("OK")],
]

event, values = sg.Window('Pick a File', layout, size=(600,100)).read(close=True)

filename = values['-IN-']
print(filename)
a=1
