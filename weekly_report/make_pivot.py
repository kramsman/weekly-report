import pandas as pd
from loguru import logger


def make_pivot(*, writer, df, report_var, index_vars, aggfunc, sheet_name, freeze_cell):
    """ template for pivot tables in reports

    Args:
        * ():
    """
    df_pt = pd.pivot_table(df, columns=report_var, index=index_vars,
                           values=['addresses_count'], margins=True, dropna=True, aggfunc=aggfunc)
    df_pt = df_pt.sort_index(axis=1, ascending=False)
    if sheet_name == '':
        sheet_name = "No Team"
    logger.debug(f"{sheet_name=}")
    df_pt.to_excel(writer, sheet_name=sheet_name, startrow=6)
    ws = writer.sheets[sheet_name]
    # ws.freeze_panes(freeze_row, freeze_col)
    mycell = ws[freeze_cell]
    ws.freeze_panes = mycell
