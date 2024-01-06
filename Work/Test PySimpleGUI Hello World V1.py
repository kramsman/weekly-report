"""
simple program to see if Gideon can run off dropbox and what needs to be done to refresh
"""

import os
import PySimpleGUI as sg

start_dir = "~/Downloads/"


print("Hello world")


layout = [
    [sg.T("Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'")],
    [sg.Text("Choose a file: "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(initial_folder= os.path.expanduser(start_dir))
    ],
    [sg.Button("OK")],
]

event, values = sg.Window('Pick a File', layout, size=(600,100)).read(close=True)

filename = values['-IN-']
print(filename)
a=1
