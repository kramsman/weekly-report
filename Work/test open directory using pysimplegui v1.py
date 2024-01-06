"""
use pysimplegui as an alternative to tkinter askopenfilename and askdirectory
"""
import os
import pathlib

# https://stackoverflow.com/questions/69491060/file-browser-gui-to-read-the-filepath

import PySimpleGUI as sg

#
# layout = [
#     [sg.T("Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'")],
#     [sg.Text("Choose a folder: "),
#      sg.Input(key="-IN-"),
#      # sg.FileBrowse(initial_folder= "~/Downloads/")
#      sg.FolderBrowse(initial_folder= os.path.expanduser("~/Downloads/"))
#     ],
#     [sg.Button("Submit")],
# ]

layout = [
    [sg.T("Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'")],
    [sg.Text("Choose a folder: ", size=(15, 1)),
      sg.Input(key='-IN-', expand_x=True),
     # sg.FileBrowse(initial_folder= "~/Downloads/")
     sg.FolderBrowse(initial_folder= os.path.expanduser("~/Library/CloudStorage/Dropbox/Postcard Files/VL Org Reports/"))
    ],
    # [sg.Text('You Entered '), sg.Text(key='-OUT-')],
    [sg.Button("Submit")],
]

window = sg.Window('My File Browser', layout, size=(1000,150))

filename = ""
while True:
    event, values = window.read()
    # values['-OUT-'] = values['-IN-']
    # window['-OUT-'].update(values['-IN-'])
    # event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    elif event == "Submit":
        filename = values['-IN-']
        break

window.close()
print(filename)
a=1
