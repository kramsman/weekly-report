"""
Create pivot table with subtotal.
Runs on report file from Sincere, prompted to input.
Now with git, generic name that replaces those with version numbers.
"""

from pathlib import Path
import pandas as pd
from openpyxl.styles import Font

from bekutils import sumby_w_totals
from bekutils import setup_loguru
from bekutils import exe_path
from bekutils import autosize_xls_cols
from bekutils import exit_yes
from bekutils import get_file_name


SINCERE_DOWNLOAD_DIR = "~/Downloads/"

exe_path = exe_path()
setup_loguru("DEBUG", "DEBUG")


def main():

    if False:
        input_file = get_file_name("Pick File",
                                   "Pick a parent-campaign file to summarize (eg "
                                   "'all-parent-campaigns-requests-2023-08-03.csv'",
                                   SINCERE_DOWNLOAD_DIR)
    else:
        # input_file = "/Users/Denise/Downloads/parent-campaign-address-counts-2023-08-03.csv"
        input_file = "/Users/Denise/Downloads/all-parent-campaigns-requests-2024-02-08.csv"

    # IF WANTED check for certain input file names after selected
    if True:
        if 'all-parent-campaigns-requests' not in str(input_file):
            exit_yes("File must have 'all-parent-campaigns-requests' in the name", "WRONG FILE TYPE")

    sincere_data = pd.read_csv(input_file)

    # FOR TESTING select only certain records
    # sincere_data = sincere_data.loc[sincere_data['Factory'] == "VA General BIPOC 7-2023"]
    # sincere_data = sincere_data[~sincere_data['Name'].str.lower().str.contains("test")]
    # sincere_data = sincere_data[sincere_data['organization'].str.lower().isin(["fl - entire state", "general",
    #                                                                            "national-bob haar"])]
    # sincere_data = sincere_data[~sincere_data['Name'].str.lower().str.contains("test")]
    # sincere_data = sincere_data[(sincere_data['Factory'].notna() & ~sincere_data['Name'].str.lower().str.contains("test"))]

    # format of INDEX_VARS_W_SUMFLAG is list of tuples: [(variable name, whether to subtotal), (var2, flag2)...]
    # Order of variables is order/level of subtotaling
    # fields to sum by. second parm is whether to total or not.
    index_vars_w_sumflag = [('factory_name', True), ('parent_campaign_name', True), ]

    # list of fields to sum/count
    sum_fields = ['addresses_count']

    ## DO THE SUMMARY
    df_pt = sumby_w_totals(sincere_data, index_vars_w_sumflag, sum_fields, 'sum')

    # op_file =op_file
    op_file = Path(__file__).stem
    writer = pd.ExcelWriter(op_file + ".xlsx")

    df_pt.to_excel(writer, sheet_name="Summary Report", startrow=6)
    wb = writer.book
    for sh in wb.worksheets:
        autosize_xls_cols(sh)

    for sh in wb.worksheets:
        sh['A1'].value = "Summary Report"
        sh['A1'].font = Font(b=True, size=20)
        sh['A3'].value = f"Source data: {input_file}"
        sh['A3'].font = Font(size=12)

    writer.close()


if __name__ == '__main__':
    main()
