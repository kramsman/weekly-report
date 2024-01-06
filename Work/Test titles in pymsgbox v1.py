""" explore why titles aren't showing up in pymsgbox with python / pymsgbox update """

# from tkinter.filedialog import askdirectory
# from tkinter.filedialog import askopenfilename
# Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

import tkinter.filedialog

print("tkinter.TkVersion=", tkinter.TkVersion)
file_name = tkinter.filedialog.askopenfilename(title="Select file")

dir_name = tkinter.filedialog.askdirectory(title="Select directory")

a=1
