"""Build pivot table sheets in Excel workbooks from Sincere request data."""

import pandas as pd
from loguru import logger


def make_pivot(*, writer: pd.ExcelWriter, df: pd.DataFrame, report_var: list[str], index_vars: list[str], aggfunc: str, sheet_name: str, freeze_cell: str) -> None:
    """Write a pivot table to an Excel sheet and freeze panes.

    Creates a pandas pivot_table with margins, sorts columns in descending
    order, writes it to the named sheet starting at row 7, and freezes panes
    at the specified cell.

    Args:
        writer: The active Excel writer to write the sheet into.
        df: Source data to pivot.
        report_var: Column(s) to use as pivot table columns (e.g. ['month']).
        index_vars: Column(s) to use as pivot table rows.
        aggfunc: Aggregation function string, e.g. 'sum' or 'count'.
        sheet_name: Target Excel sheet name; substituted with 'No Team' if
            empty.
        freeze_cell: Excel cell reference at which to freeze panes, e.g.
            'C10'.
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
