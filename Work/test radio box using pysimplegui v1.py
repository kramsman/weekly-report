"""
use pysimplegui as an alternative to tkinter askopenfilename and askdirectory
USE THIS VERSION - V2 shortens one shot window to a few lines
"""

import PySimpleGUI as sg

def bek_radio_box_w_ok(box_title, title2, radio_choices, choice_text=None):
    if choice_text is None:
        choice_text = True
    else:
        choice_text = False

    layout = [
                [sg.Text(title2)],
                [sg.Radio(text, 0, default=(True if idx == 0 else False)) for idx, text in enumerate(radio_choices)],
                [sg.Button('Ok', focus = True)]
             ]

    event, values = sg.Window(box_title, layout, font=("Arial", 16) ).read(close=True)

    for idx, val in values.items():
        if val:  # the first (should be only one) chosen
            if choice_text:
                return_val = radio_choices[idx].lower()
            else:
                return_val = idx
            break

    return return_val


def radio_box(box_title, title2, radio_choices, choice_text=None):
    if choice_text is None:
        choice_text = True
    else:
        choice_text = False

    layout = [
                [sg.Text(title2)],
                [sg.Radio(text, 0, enable_events = True) for idx, text in enumerate(radio_choices)],
             ]

    event, values = sg.Window(box_title, layout, titlebar_font=("Arial", 18), font=("Arial", 16),
                              use_custom_titlebar=True).read(close=True)

    for idx, val in values.items():
        if val:  # item clicked, the first (should be only one) chosen
            if choice_text:
                return_val = radio_choices[idx].lower()
            else:
                return_val = idx
            break

    return return_val

pick = radio_box("Pick a Value", "Line one\nLine 2", ["Run", "Update", "Exit"])
print(f"{pick=}")
pick = bek_radio_box_w_ok("Pick a Value", "Default is first choice", ["Run", "Update", "Exit"])
print(f"{pick=}")

a=1
