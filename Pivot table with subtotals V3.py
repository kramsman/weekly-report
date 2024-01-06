""" create pivot table with subtotal """
import pathlib

# V2 play with total in multiindex
# V2.1 clean up commented lines; test replacing multiindex columns by position
# V3 specify fields to have total lines

import pandas as pd
import numpy as np
import PySimpleGUI as sg
import os
import inspect
from openpyxl.styles import Font
import itertools

SINCERE_DOWNLOAD_DIR = "~/Downloads/"
TOTAL_STR = '_TOTAL'
index_vars_w_sumflag = [('Factory', True), 'Name']


def autosizexls(ws):
    # from https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size
    # print(ws.title)
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                # dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
                # print("cell value", cell.value, "data_type",cell.data_type)
                if(cell.data_type == 'd'):
                    dateWidth = 10
                else: dateWidth = len(str(cell.value))
                # dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), dateWidth))

    for col, value in dims.items():
        # print(col , value)
        ws.column_dimensions[col].width = value + 1

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

    col_factor = 3  # to scale window equally
    row_factor = 25  # to scale window equally
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


def main(index_vars_w_sumflag):

    # /Users/Denise/Downloads/
    #     input_file = get_file_name("Pick a File",
    #                                "Select Sincere address export file 'parent-campaign-address-counts-yyyy-mm-dd.csv'",
    #                                SINCERE_DOWNLOAD_DIR)

    # todo: force index to str so we can concat with '_sum'

    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 120)
    pd.options.display.multi_sparse = False

    # index_vars_w_sumflag = [('Parent Campaign', True), 'Child Organization', ('Name', True)]
    # index_vars_w_sumflag = [('Factory', False), ('Name', True)]
    # index_vars_w_sumflag = [('Factory', False), 'Name']
    index_vars_w_sumflag2 = [val if isinstance(val, tuple) else (val, False)
                             for val in index_vars_w_sumflag]
    index_var_dict = {val[0]: val[1] for val in index_vars_w_sumflag2}
    index_vars = list(index_var_dict.keys())
    # sum_vars = [tuple[0] for index, tuple in enumerate(index_vars_w_sumflag2) if tuple[1]]
    sum_vars = [field for index, (field, sum_flag) in enumerate(index_vars_w_sumflag2) if sum_flag]
    # sum_cols = [index for index, tuple in enumerate(index_vars_w_sumflag2) if tuple[1]]
    # index_vars = ['Break_1', 'Factory', 'Name']

    input_file = get_file_name("Pick File","Pick a parent-campaign file to summarize (eg "
                                           "'parent-campaign-address-counts-2023-08-03.csv'", SINCERE_DOWNLOAD_DIR)
    # input_file = "/Users/Denise/Downloads/parent-campaign-address-counts-2023-08-03.csv"
    if 'parent-campaign-address-counts' not in input_file:
        exit_yes("File must have 'parent-campaign-address-counts' in the name", "WRONG FILE TYPE")
    df_combine = pd.read_csv(input_file)

    # select only certain records
    # df_combine = df_combine.loc[df_combine['Factory'] == "VA General BIPOC 7-2023"]
    df_combine = df_combine[~df_combine['Name'].str.lower().str.contains("test")]

    df_combine['Remaining In Room'] = df_combine['Assigned to Organizations'] - df_combine['Assigned to Writers']
    # writer = pd.ExcelWriter("Pivot with subtotals.xlsx", engine="openpyxl")

    # df_base = df_combine.groupby(index_vars).agg({"Total Addresses":'sum','Available Addresses':'sum'})
    flds = ['Total Addresses', 'Available Addresses', 'Assigned to Organizations', 'Assigned to Writers', 'Remaining In Room']
    # dictionary of all summed fields field:sum
    flds_dict = {fld:sum for fld in flds}

    # create df summed by break of all fields in index_vars
    df_base = df_combine.groupby(index_vars).agg(flds_dict)

    # df_base = df_combine.groupby(index_vars).agg({'Total Addresses':sum, 'Available Addresses':sum,
    #                                'Assigned to Organizations':sum, 'Assigned to Writers':sum})
    print(f"{df_base=}")

    # fds sill be alist of df objects, each a summ on a different level
    dfs = []

    # series1 is grand total for all
    series1 = df_base.sum()
    dfs.append(series1.to_frame().transpose())

    print('combinations')
    for L in range(1,len(sum_vars) + 1):
        for temp_sum_vars in itertools.combinations(sum_vars, L):
            print(temp_sum_vars)
            dfs.append(df_base.groupby(level=temp_sum_vars).sum())

    # get index into correct format with correct number of fields (len(index_vars)) and TOTAL_STR in correct columns
    for df in dfs:
        # check index type because some are not multiindex and have no len (grand total), some str and not iterable
        index_obj = df.index.values[0]  # so we can check index type
        if type(df.index.values[0]) == np.int64:  # grand total row
            index_array = len(index_vars) * [[TOTAL_STR]]
            df.index = pd.MultiIndex.from_arrays(index_array)
        elif isinstance(index_obj, (list, tuple)):
            index_array = len(index_vars) * [len(df.index.values)*[TOTAL_STR]]
            # this is a neat trick: list(zip(*dfs[2].index.values)) unzips and produces a list of lists (actually
            # tuples), each the n position in all tuples of index.
            # https://stackoverflow.com/questions/12974474/how-to-unzip-a-list-of-tuples-into-individual-lists
            untupled_indices = list(zip(*df.index.values))
            field_w_index_list = list(zip(df.index.names, untupled_indices))  # tuples of col name and list
            # of index values
            field_to_index_dict = {field_and_index[0]: field_and_index[1] for field_and_index in field_w_index_list}
            for column_field in df.index.names:
                col_num = [index_vars_w_sumflag[0] for index_vars_w_sumflag in index_vars_w_sumflag2].index(column_field)
                index_array[col_num] = field_to_index_dict[column_field]
            df.index = pd.MultiIndex.from_arrays(index_array)
        # elif isinstance(index_obj, str):
        elif len(df.index.names) == 1:
            index_array = len(index_vars) * [len(df.index.values)*[TOTAL_STR]]

            # sub index_var position below
            index_array[index_vars.index(df.index.names[0])] = df.index.values
            df.index = pd.MultiIndex.from_arrays(index_array)
        else:
            print(df.index.values[0], ' is not tuple or list or np.int64')
            index_array = df.index.values + \
                          (len(index_vars) - len(df.index.values[0])) * [[TOTAL_STR]]
            df.index = pd.MultiIndex.from_arrays(index_array)

    # concat all dataframes together, sort index
    df_combine = df_base
    for df in dfs:
        df_combine = pd.concat([df_combine, df])

    df_combine.index.names = index_vars
    level_list = list(range(0, len(index_vars)-1))

    df_combine = df_combine.sort_index(level=level_list)
    print('df_combine index names', df_combine.index.names)
    # print(df_combine.index.values)
    print(df_combine)

    # op_file =op_file
    op_file = pathlib.Path(__file__).stem
    writer = pd.ExcelWriter(op_file + ".xlsx")

    df_combine.to_excel(writer, sheet_name="Summary Report", startrow=6)
    wb = writer.book
    for sh in wb.worksheets:
        autosizexls(sh)

    for sh in wb.worksheets:
        sh['A1'].value = "Summary Report"
        sh['A1'].font = Font(b=True, size=20)
        sh['A3'].value = "Source data: " + input_file
        sh['A3'].font = Font(size=12)

    writer.close()

    a = 1


main(index_vars_w_sumflag)
a = 1
