import inspect

from loguru import logger
from uvbekutils import exit_yes
from uvbekutils import exit_yes


def check_sheet_headers(ws, vals):
    """
    Check list of (cell, val) tuples representing header labels in ws_to_chk and error if val not found in cell.
    eg vals = [('A1', 'use'), ('B1', 'fromFilePath'), ('C1', 'fromfilename'), ....]
    """

    def chk_header_vals(ws_to_chk, cell, val):
        """ error if val not found in wks cell. """
        if str(ws_to_chk[cell].value).strip().lower() != str(val).lower():
            exit_yes((f"Column heading '{cell}' on Setup sheet '{ws_to_chk.title}' not equal to literal '{val}'."
                      f"\n\nIt is '{str(ws_to_chk[cell].value)}'."),
                     )

    for pairs in vals:
        chk_header_vals(ws, pairs[0], pairs[1])


def check_df_headers(df, vals):
    """ Check list of values and compare to df column names.  Retun False if not matched. """

    vals = [val.lower().strip() for val in vals]
    cols = [col.lower().strip() for col in df.columns]

    if vals != cols:
        exit_yes((f"Dataframe columns do not match checked values.\n\n"
                  f"Columns are:\n\n"
                  f"'{','.join(cols)}\n\n"
                  f"Checked vals are:\n\n"
                  f"'{','.join(vals)}\n\n"
                  ),
                 )


def calling_func(level=0):
    """ returns the various levels of calling function.  0 is current, 1 is caller of current, etc """
    try:
        func = f"'{inspect.stack()[level][3]}', line #: {inspect.stack()[level][2]}"
    except Exception:
        func = f"** error ** inspect level too deep: {str(level)} called from {inspect.stack()[level][3]}"
    return func


def df_to_sheet(writer, df, sheet_name, freeze_cell=None):
    """ write df info to excel writer and freeze if specified """
    if sheet_name == '':
        sheet_name = "No Team"
    logger.debug(f"{sheet_name=}")
    df.to_excel(writer, sheet_name=sheet_name, startrow=6)
    ws = writer.sheets[sheet_name]
    # ws.freeze_panes(freeze_row, freeze_col)
    if freeze_cell is not None:
        mycell = ws[freeze_cell]
        ws.freeze_panes = mycell


def max_used_col(ws, row):
    """ Returns the column number (1 indexed) maximum non-none column in the input row of a sheet. """
    mxcol = 0
    for cell in reversed(ws[row]):
        if cell.value is not None:
            mxcol = cell.col_idx
            break
    return mxcol


def max_used_row(ws, col):
    """ Returns the column number (1 indexed) maximum non-none column in the input row of a sheet. """
    # for max_row, row in enumerate(ws, 1):
    #     if all(c.value is None for c in row):
    #         break
    mxrow = 0
    for cell in reversed(ws[col]):
        if cell.value is not None:
            mxrow = cell.row_idx
            break
    return mxrow
