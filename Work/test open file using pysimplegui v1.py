"""
use pysimplegui as an alternative to tkinter askopenfilename and askdirectory
"""
import os
import pathlib

# https://stackoverflow.com/questions/69491060/file-browser-gui-to-read-the-filepath

import PySimpleGUI as sg

# left_col = [[sg.Text('Folder'), sg.In(size=(25, 1), enable_events=True, key='-FOLDER-'), sg.FolderBrowse()]]
# layout = [[sg.Column(left_col, element_justification='c')]
#           window = sg.Window('Multiple Format Image Viewer', layout, resizable=True)
#
# while True:
#     event, values = window.read()
#     if event in (sg.WIN_CLOSED, 'Exit'):
#         break
#     if event == '-FOLDER-':
#         folder = values['-FOLDER-']


layout = [
    [sg.Text("Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'")],
    [sg.Text("Choose a file: "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(initial_folder= os.path.expanduser("~/Downloads/"))
    ],
    [sg.Button("Submit")],
]

window = sg.Window('My File Browser', layout, size=(600,150))

filename = ""
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    elif event == "Submit":
        filename = values['-IN-']
        break

window.close()
print(filename)
a=1
