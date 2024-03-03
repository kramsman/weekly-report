"""
Create pivot table with subtotal.
Runs on report file from Sincere, prompted to input.
Now with git, generic name that replaces those with version numbers.
"""

# V2 play with total in multiindex
# V2.1 clean up commented lines; test replacing multiindex columns by position
# V3 specify fields to have total lines
# Started as V3

# import pathlib
from pathlib import Path

# TODO collapse all parent lines for factory locked

import openpyxl.cell.cell
import pandas as pd
# import numpy as np
# import PySimpleGUI as sg
# import os
# import sys
# import inspect
from openpyxl.styles import Font
# import itertools
# from loguru import logger

# from bekutils.sumby_w_totals import sumby_w_totals
from bekutils import sumby_w_totals
from bekutils import setup_loguru
from bekutils import exe_file
# from bekutils import autosize_xls_cols
# from bekutils import text_box
from bekutils import exit_yes
# from bekutils import exit_yes_no
from bekutils import get_file_name
# from bekutils import get_dir_name
from bekutils import bek_excel_titles


# format of INDEX_VARS_W_SUMFLAG is list of tuples: (variable name, whether to subtotal)
# Order of variables is order/level of subtotaling

SINCERE_DOWNLOAD_DIR = "~/Downloads/"

# exe_path = exe_path()

setup_loguru("DEBUG", "DEBUG")


# def bek_excel_titles(wb, sheet_name_list, cell_infos = None, auto_size_before=None, auto_size_after=None, close_wb=None):
#     """
#     Write titles to an Excel file.  Can autosize before or after titles are inserted (usually before because titles
#     are wide) to avoid ##### in cells.
#     Args:
#         sheet_name_list (): list of sheet names to add titles. 'True' will put titles on all.
#         cell_infos (): list of cell attributes used to format specified as dictionaries. Rows and cols 1 based
#         numerics. 'value's are replaced as text; font is passed as eval of cell value (note in function):
#             {'row':1,'col':1, 'cell_attr':"value", 'cell_value':'Summary Report'},
#             {'row':1,'col':1, 'cell_attr':"font", 'cell_value':'Font(b=True, size=20)'},
#     """
#
#     # from pathlib import Path
#     # import pandas as pd
#     from openpyxl.styles import Font
#     from bekutils import autosize_xls_cols
#     # from bekutils import exe_file
#
#     # op_file = f"{Path(__file__).stem}.xlsx"
#     # op_file = exe_file().with_suffix(".xlsx")
#
#     # writer = pd.ExcelWriter(op_file)
#
#     # df.to_excel(writer, sheet_name = sheet_name, startrow = startrow)
#     # wb = writer.book
#
#     if isinstance(sheet_name_list,str):
#         sheet_name_list = [sheet_name_list]
#     elif isinstance(sheet_name_list, list):
#         pass
#     else:
#         exit_yes(f"sheet_name_list is not str or list: {sheet_name_list=}", "Error", raise_err=True)
#
#     if auto_size_before:
#         for sh in wb.worksheets:
#             autosize_xls_cols(sh)
#
#     # TODO perform checks on formats of cell_infos
#     if cell_infos:
#         for sh in wb.worksheets:
#             if sh.title in sheet_name_list:
#                 for cell_info in cell_infos:
#                     if cell_info['cell_attr'] == 'font':
#                         # use eval because needs to be like ft1 = Font(name='Arial', size=14).  might be able to use
#                         # another setattr but not ready to try now
#                         setattr(sh.cell(row=cell_info['row'], column=cell_info['col']), cell_info['cell_attr'],
#                                 eval(cell_info['cell_value']))
#                     else:
#                         setattr(sh.cell(row=cell_info['row'], column=cell_info['col']), cell_info['cell_attr'],
#                                 cell_info['cell_value'])
#
#     if auto_size_after:
#         for sh in wb.worksheets:
#             autosize_xls_cols(sh)
#



def main():
    # /Users/Denise/Downloads/
    #     input_file = get_file_name("Pick a File",
    #                                "Select Sincere address export file 'parent-campaign-address-counts-yyyy-mm-dd.csv'",
    #                                SINCERE_DOWNLOAD_DIR)

    # initialize_logger(ROOT_PATH)

    # todo: force index to str so we can concat with '_sum'

    # pd.set_option('display.max_columns', None)
    # pd.set_option('display.width', 120)
    # pd.options.display.multi_sparse = False

    if False:
        input_file = get_file_name("Pick File",
                                   f"Pick a parent-campaign file to summarize (eg "
                                   f"'parent-campaign-address-counts-2023-08-03.csv'."
                                   f"\n\nCreate via ROV > Reports > New Report > Parent Campaign Address Counts",
                                   SINCERE_DOWNLOAD_DIR)
    else:
        input_file = "/Users/Denise/Downloads/parent-campaign-address-counts-2024-02-26.csv"
        # input_file = "/Users/Denise/Downloads/all-users-2024-01-06.csv"

    if False:
        if 'parent-campaign-address-counts' not in str(input_file):
            exit_yes("File must have 'parent-campaign-address-counts' in the name", "WRONG FILE TYPE")

    sincere_data = pd.read_csv(input_file)
    # print(sincere_data.dtypes)
    # aaa = sincere_data.loc[sincere_data['Factory'].isna()]

    # FOR TESTING select only certain records
    # sincere_data = sincere_data.loc[sincere_data['Factory'] == "VA General BIPOC 7-2023"]
    # sincere_data = sincere_data[~sincere_data['Name'].str.lower().str.contains("test")]
    # sincere_data = sincere_data[sincere_data['organization'].str.lower().isin(["fl - entire state", "general",
    #                                                                            "national-bob haar"])]
    a = 1

    sincere_data = sincere_data[(sincere_data['Factory'].notna() & ~sincere_data['Name'].str.lower().str.contains("test"))]

    # sincere_data = sincere_data[~sincere_data['Name'].str.lower().str.contains("test")]
    sincere_data = sincere_data[(~sincere_data['Factory'].str.lower().str.contains("zzz"))]

    # limit to only this year's campaigns
    sincere_data = sincere_data[(sincere_data['Name'].str.lower().str.contains("2024"))]


    sincere_data['Remaining In Room'] = sincere_data['Assigned to Organizations'] - sincere_data['Assigned to Writers']

    # collapse parent campaigns to one line for locked Factories
    sincere_data.loc[sincere_data['Is Locked'] == True, 'Name'] = 'Currently Locked-All Parent Campaigns'

    # fields to sum by. second parm is whether to total or not.
    index_vars_w_sumflag = [('Factory', True), ('Name', False), ]
    # index_vars_w_sumflag = [('organization', True), ('team', True),]

    # list of fields to sum/count
    sum_fields = ['Total Addresses', 'Available Addresses', 'Assigned to Organizations', 'Assigned to Writers',
                     'Remaining In Room']
    # sum_fields = ['email']
    df_pt = sumby_w_totals(sincere_data, index_vars_w_sumflag, sum_fields, 'sum')

    # calc % assigned to rooms
    df_pt['Percent Assigned to Rooms'] = round(100 * df_pt['Assigned to Organizations'] / df_pt['Total Addresses'], 0)

    # rename 'Available Addresses' to 'Available For Rooms' to avoid confusion with number available for rooms vs in
    # rooms.
    df_pt = df_pt.rename(columns={'Available Addresses': 'Available For Rooms', 'Assigned to Organizations':
        'Assigned to Rooms'})

    op_file = exe_file().with_suffix(".xlsx")
    writer = pd.ExcelWriter(op_file)
    df_pt.to_excel(writer, sheet_name = "test", startrow = 6)

    # bek_write_excel(df_pt, sheet_name = "Summary Report", startrow=6,
    #                 cell_infos=[{'row':1,'col':1, 'cell_attr':"value", 'cell_value':'Summary Report'},
    #                             {'row':1,'col':1, 'cell_attr':"font", 'cell_value':'Font(b=True, size=20)'},
    #                             {'row':3,'col':1, 'cell_attr':"value", 'cell_value':f"Input file: {input_file}"},
    #                             {'row':3,'col':1, 'cell_attr':"font", 'cell_value':'Font(size=12)'},
    #                             ],
    #                 )

    bek_excel_titles(writer.book, sheet_name_list = "test", auto_size_before=True,
                    cell_infos=[{'row':1,'col':1, 'cell_attr':"value", 'cell_value':'Campaign Summary Report'},
                                {'row':1,'col':1, 'cell_attr':"font", 'cell_value':'Font(b=True, size=20)'},
                                {'row': 2, 'col': 1, 'cell_attr': "value", 'cell_value': 'The Center for Common '
                                                                                         'Ground'},
                                {'row': 2, 'col': 1, 'cell_attr': "font", 'cell_value': 'Font(b=True, size=12)'},
                                {'row':3,'col':1, 'cell_attr':"value", 'cell_value':f"Input file: {input_file}"},
                                {'row':3,'col':1, 'cell_attr':"font", 'cell_value':'Font(size=12)'},
                                ],
                    )

    a=1
    writer.close()

    # sh['A1'].value = "Summary Report"
    # sh['A1'].font = Font(b=True, size=20)
    # sh['A3'].value = f"Source data: {input_file}"
    # sh['A3'].font = Font(size=12)

    # cell_infos = [4, 1, "value", 'Summary Report'), ],
    # cell_infos = ["sh['A1'].value = 'Summary Report'",
    #               f"sh['A3'].value = 'Source data:' <data here> ", ],
    #
    # # op_file =op_file
    # op_file = Path(__file__).stem
    # writer = pd.ExcelWriter(op_file + ".xlsx")
    #
    # df_pt.to_excel(writer, sheet_name="Summary Report", startrow=6)
    # wb = writer.book
    # for sh in wb.worksheets:
    #     autosize_xls_cols(sh)
    #
    # for sh in wb.worksheets:
    #     sh['A1'].value = "Summary Report"
    #     sh['A1'].font = Font(b=True, size=20)
    #     sh['A3'].value = f"Source data: {input_file}"
    #     sh['A3'].font = Font(size=12)
    #
    # writer.close()

    a = 1


if __name__ == '__main__':
    main()
    a = 1
