""" create pivot table with subtotal """
# V2 play with total in multiindex
# V2.1 clean up commented lines; test replacing multiindex columns by position

import pandas as pd
import numpy as np
import PySimpleGUI as sg
import os
import inspect

SINCERE_DOWNLOAD_DIR = "~/Downloads/"
TOTAL_STR = '_TOTAL'

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

    col_factor = 3  # to scale windo equally
    row_factor = 25  # to scale windo equally
    max_cols = len(max(txt.split("\n"), key=len)) * col_factor
    cols = max_cols
    # v_scroll = False
    col_limit = 80 * col_factor
    col_min = 50 * col_factor
    if cols > col_limit:
        # v_scroll = True
        cols = col_limit
    elif cols < col_min:
        cols = col_min

    h_scroll = False
    row_limit = 80
    row_min = 10
    max_rows = len(txt.split("\n"))
    rows = max_rows
    if rows > row_limit:
        h_scroll = True
        rows = row_limit
    elif rows < row_min:
        rows = row_min

    layout = [
        [sg.Text(title2, font=("Arial", 18))],
        [sg.Multiline(txt, autoscroll=False, horizontal_scroll=h_scroll, expand_x=True,
                      expand_y=True, enable_events=True)],
        [sg.Button(text) for text in buttons],
    ]

    event, values = sg.Window(box_title, layout, titlebar_font=("Arial", 20), font=("Arial", 14),
                              use_custom_titlebar=True, size=(600, rows * row_factor), disable_close=True,
                              resizable=True, grab_anywhere=True).read(close=True)
    # event, values = sg.Window(box_title, layout, titlebar_font=("Arial", 20), font=("Arial", 14),
    #                           use_custom_titlebar=True, size=(cols,rows), disable_close=True,
    #                           resizable=True, grab_anywhere	= True).read(close=True)
    if event is not None:
        event = event.lower()
    return event


def exit_yes(msg: str, title: str = None, *, errmsg: str = None) -> None:
    """ exits program after giving user a popup window and raising an error. """
    msg = (msg + "\n\n\nExiting." +
           f"\n\nCalled from {calling_func(level=3)}"
           f"\nCalled from {calling_func(level=2)}"
           f"\nCalled from {calling_func(level=1)}"
           )
    if not errmsg:
        errmsg = msg.replace("\n", " ")  # dont fill the console with linefeeds
    if not title:
        title = "** Exiting Program **"
    # pymsgbox.alert(msg, title)
    bek_text_box(title, "", "\n\n" + msg, ["Ok"])
    raise Exception(errmsg)


def exit_yes_no(msg, title=None, display_exiting=False):
    """ makes this choice to continue one line"""
    if not title:
        title = "Exit?"
    # choice = pymsgbox.confirm(msg, title, ['Yes', 'No'])
    choice = bek_text_box(title, "", "\n\n" + msg, ['Yes', 'No'])

    if choice == "no":
        if display_exiting:
            # pymsgbox.alert("Exiting", "Alert")
            bek_text_box("Alert", "", "\n\nExiting", ["Ok"])
        exit()


def calling_func(level=0):
    """ returns the various levels of calling function.  0 is current, 1 is caller of current, etc """
    try:
        func = f"'{inspect.stack()[level][3]}', line #: {inspect.stack()[level][2]}"
    except Exception:
        func = f"** error ** inspect level too deep: {str(level)} called from {inspect.stack()[level][3]}"
    return func


def get_file_name(box_title, title2, initial_dir):
    """ show an "Open" dialog box and return the selected file name. Replaced askopenfilename with pyeasygui
    :param title2: heading of the box
    :type title2: text next to input field
    """
    # "Select Sincere address export file 'all-parent-campaign-requests-yyyy-mm-dd.csv'"
    layout = [
        [sg.Text(title2, font=("Arial", 18))],
        [
            sg.Input(key="-IN-", expand_x=True),
            sg.FileBrowse(initial_folder=os.path.expanduser(initial_dir))
        ],
        [sg.Button("Choose")],
    ]

    # event, values = sg.Window(heading_in_box, layout, size=(600, 100)).read(close=True)
    event, values = sg.Window(box_title, layout, titlebar_font=("Arial", 20), font=("Arial", 14),
                              size=(1000, 150), use_custom_titlebar=True).read(close=True)
    # sg.Window.close()

    file_name = values['-IN-']
    if file_name == "":
        exit_yes("No file name chosen")

    return file_name


def make_pivot(writer, df, report_var, index_vars, aggfunc, sheet_name, freeze_cell):
    """ template for pivot tables in reports """
    df_pt = pd.pivot_table(df, columns=report_var, index=index_vars,
                           values=['addresses_count'], margins=True, dropna=True, aggfunc=aggfunc)
    df_pt = df_pt.sort_index(axis=1, ascending=False)
    df_pt.to_excel(writer, sheet_name=sheet_name, startrow=6)
    ws = writer.sheets[sheet_name]
    # ws.freeze_panes(freeze_row, freeze_col)
    mycell = ws[freeze_cell]
    ws.freeze_panes = mycell

    return df_pt


def main():
    # /Users/Denise/Downloads/
    #     input_file = get_file_name("Pick a File",
    #                                "Select Sincere address export file 'parent-campaign-address-counts-yyyy-mm-dd.csv'",
    #                                SINCERE_DOWNLOAD_DIR)

    # todo: force index to str so we can concat with '_sum'

    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 120)
    index_vars = ['Break_1', 'Factory', 'Name']

    input_file = '/Users/Denise/Downloads/TEST parent-campaign-address-counts-2023-05-22.csv'
    df_combine = pd.read_csv(input_file)


    writer = pd.ExcelWriter("Pivot with subtotals.xlsx", engine="openpyxl")

    df_base = pd.pivot_table(df_combine,  # columns=['Amount'],
                           #    index=['Break_1', 'Factory', 'Name'],
                           index=index_vars,
                           values=['Total Addresses', 'Available Addresses', 'Assigned to Organizations', 'Assigned to '
                                                                                                          'Writers'],
                           fill_value=0, aggfunc=np.sum, dropna=True, )
    print('df_base',df_base)

    dfs = {}

    series1 = df_base.sum()
    dfs[0] = series1.to_frame().transpose()
    print('df0',dfs[0])
    print(dfs[0].index.names)
    print(dfs[0].index.values)

    for df_num in range(1,len(index_vars)):
        level_list = list(range(0,df_num))
        dfs[df_num] = df_base.groupby(level=level_list).sum()
        # print(f"{df_num=}",dfs[df_num])
        # print(dfs[df_num].index.names)
        # print(dfs[df_num].index.values)

    # now set multiindex - cols up to df_num blank; df_num appends _sum to vals; others get _TOTAL
    # len(dfs[2].index.values[0]) gives # of elements in each tuple
    # type(dfs[0].index.values[0]) is <class 'pandas.core.indexes.range.RangeIndex'> <class 'numpy.int64'> <class
    # 'pandas.core.indexes.multi.MultiIndex'>
    # for dfs[0]; 1 for first (XX), 2 for 2nd (XX/AA)
    for df_num in range(0, len(index_vars)):
        # index_array = []
        index_array = len(index_vars) * [["dummy"]]  # initialize a list of list items
        index_array[1] = [1,2,3]
        # check index type because some are not multiindex and have no len (grand total), some str and not iterable
        index_obj = dfs[df_num].index.values[0]
        if isinstance(index_obj, (list, tuple)):
            # this is a neat trick: list(zip(*dfs[2].index.values)) unzips and produces a list of lists (actually
            # tuples), each the n position in all tuples of index.
            # https://stackoverflow.com/questions/12974474/how-to-unzip-a-list-of-tuples-into-individual-lists
            untupled_indices = list(zip(*dfs[df_num].index.values))
            index_array2 = (len(index_vars) - len(dfs[df_num].index.values[0])) \
                           * [len(dfs[df_num].index.values)*[TOTAL_STR]]
            index_array = untupled_indices + index_array2
            dfs[df_num].index = pd.MultiIndex.from_arrays(index_array)

        elif isinstance(index_obj, str):
            index_array1 = [list(dfs[df_num].index.values)]
            index_array2 = (len(index_vars) - 1) * [len(dfs[df_num].index.values)*[TOTAL_STR]]
            index_array = index_array1 + index_array2
            dfs[df_num].index = pd.MultiIndex.from_arrays(index_array)
        else:
            print(dfs[df_num].index.values[0], ' is not tuple or list')
            if type(dfs[df_num].index.values[0]) == np.int64:  # grand total row
                index_array = len(index_vars) * [[TOTAL_STR]]
                dfs[df_num].index = pd.MultiIndex.from_arrays(index_array)
            else:
                index_array = dfs[df_num].index.values + \
                              (len(index_vars) - len(dfs[df_num].index.values[0]) ) * [[TOTAL_STR]]
                dfs[df_num].index = pd.MultiIndex.from_arrays(index_array)

    # concat all dataframes together, sort index
    df_combine = df_base
    for df_num in range(0,len(index_vars)):
        df_combine = pd.concat([df_combine, dfs[df_num]])

    level_list = list(range(0,len(index_vars)-1))

    df_combine = df_combine.sort_index(level=level_list)
    print('df_combine', df_combine.index.names)
    print(df_combine.index.values)
    print(df_combine)

    a = 1


main()
a = 1
