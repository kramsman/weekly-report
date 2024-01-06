"""
use pysimplegui as an alternative to tkinter askopenfilename and askdirectory
USE THIS VERSION - V2 shortens one shot window to a few lines
"""
# taken from 'single shot' here: https://www.pysimplegui.org/en/latest/cookbook/

import os

import PySimpleGUI as sg

# def print_multi(self, *args, end=None, sep=None, text_color=None,
#     background_color=None, justification=None, font=None, colors=None, t=None,
#     b=None, c=None,  autoscroll=True):

        # window['-STLINE-'].print(values['word'])


def bek_text_box(box_title, title2, txt, buttons=None):
    """" Display text block with lines separated by \n and choice of buttons at bottom.
    :param box_title: main heading on box
    :type box_title: str
    :param title2: 2nd title above text
    :type title2: str
    :param txt: text block with lines separated by \n
    :type txt: str
    :param buttons: list of button text, defaults to ['OK', 'Exit']
    :type buttons: list of str
    :return: lower case value of selected button
    :rtype: str
    """

    if buttons is None:
        buttons = ["OK", "Exit"]

    col_factor = 10 # to scale windo equally
    row_factor = 10 # to scale windo equally

    max_cols = len(max(txt.split("\n"), key=len))
    cols = max_cols
    # v_scroll = False
    col_limit = 80
    col_min = 20
    if cols > col_limit:
        # v_scroll = True
        cols = col_limit
    elif cols < col_min:
        cols = col_min


    h_scroll = False
    row_limit = 80
    row_min = 20
    max_rows = len(txt.split("\n"))
    rows = max_rows
    if rows > row_limit:
        h_scroll = True
        rows = row_limit
    elif rows < row_min:
        rows = row_min

    # , justification= 'c', pad=30

    layout = [
        [sg.Text(title2,font=("Arial", 18))],
        [sg.Multiline(txt, autoscroll=False, horizontal_scroll = h_scroll, expand_x=True,
                      expand_y=True, enable_events=True )],
        [sg.Button(text) for text in buttons],
    ]

    # margins=(30,30),

    event, values = sg.Window(box_title, layout, titlebar_font=("Arial", 20), font=("Arial", 14),
                              use_custom_titlebar=True, size=(cols*col_factor,rows*row_factor),
                              disable_close=True,
                              resizable=True, grab_anywhere	= True).read(close=True)

    if event is not None:
        event = event.lower()
    return event


with open('/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/PythonProgs/VoterLetters/Small Writer report.py') as f:
    # lines = f.readlines()
    lines = f.read()


# try popup yes no for the heck of it
ch = sg.popup_scrolled(lines,  title="popup_scrolled", auto_close=True, size=(100,50), auto_close_duration=99999)
ch = sg.popup_cancel("box text here.",  title="Title Here", auto_close=True, auto_close_duration=5)
ch = sg.popup_ok_cancel("box text here.",  title="Title Here", auto_close=True, auto_close_duration=5)
ch = sg.popup_error("box text here.",  title="Title Here", auto_close=True, auto_close_duration=5)
ch = sg.popup_yes_no(lines,  title="popup_yes_no", auto_close=True, auto_close_duration=99999)


# pick = bek_text_box("Review Text", "This is first_code:", "Help me with this test.\nI relly mean it")
pick = bek_text_box("Alert", "", "\n\nAdding Directory via bad_path_create", ["OK"])
pick = bek_text_box("Review Text", "Instructions:", lines)
pick = bek_text_box("Review Text", "Instructions:", lines, ["Run", "Upload", "Exit"])

a=1
