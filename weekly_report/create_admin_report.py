"""Build the ROV-wide admin Excel workbook from Sincere request data."""

from collections.abc import Callable
from pathlib import Path

import pandas as pd
from loguru import logger
from openpyxl.styles import Font
from uvbekutils import autosize_xls_cols
from uvbekutils import exit_yes

from factory_and_campaign_subtotals import factory_and_campaign_subtotals
from weekly_report.constants import FACTORY_FILTER_STRING
from weekly_report.make_chart import make_chart
from weekly_report.make_pivot import make_pivot
from weekly_report.utils.utils import df_to_sheet


def create_admin_report(
        sincere_df: pd.DataFrame,
        sincere_data_file: str,
        report_by: str,
        str_output_dir_admin: str,
        admin_rpt_filename: str,
        factory_csv: Path,
        df_sort_func: Callable,
) -> None:
    """Create the ROV-wide admin Excel workbook with charts and pivot tables.

    Writes a daily-requests line chart, factory/campaign subtotal sheets,
    and election/master/room pivot tables to a single .xlsx file. Column
    widths are auto-sized and header titles are stamped on every sheet.

    Args:
        sincere_df: Filtered Sincere request data for all orgs and factories.
        sincere_data_file: Source CSV filename, used in sheet header notes,
            e.g. 'all-parent-campaigns-requests-2024-02-26.csv'.
        report_by: 'W' for weekly or 'M' for monthly column grouping.
        str_output_dir_admin: Directory path where the admin .xlsx is written.
        admin_rpt_filename: Output filename for the admin report .xlsx.
        factory_csv: Path to the Sincere parent-campaign-address-counts CSV
            used by factory_and_campaign_subtotals().
        df_sort_func: Callable that maps a MultiIndex level to sort keys;
            used to sort factory/election subtotal DataFrames.
    """

    print("Org: Enterprise wide")  # TODO: use logger instead of prints

    report_by = report_by.upper()
    if report_by == 'M':
        report_var = "month"
    elif report_by == 'W':
        report_var = "data_week_ending"
    else:
        exit_yes(f"Report_by not W or M: {report_by}")

    output_dir_admin = Path(str_output_dir_admin).expanduser()

    excel_output_file = Path(output_dir_admin).expanduser() / admin_rpt_filename

    writer = pd.ExcelWriter(excel_output_file, engine="openpyxl")

    # create chart of daily requests
    logger.debug("calling make_chart")
    make_chart(writer, sincere_df, 'factory_name', 'Chart')

    # Report on Masters only using sum w subtotals
    factory_tots, factory_pull_date = factory_and_campaign_subtotals(factory_csv, FACTORY_FILTER_STRING,
                                 break_fields="[('Election', True), ('Factory', False), ]")
    logger.debug("completed factory_and_campaign_subtotals")

    factory_tots.sort_index(level=['Election', 'Factory'], key=df_sort_func, ascending=True, inplace=True)
    df_to_sheet(writer, factory_tots, 'Assigned_by_state', freeze_cell="C8")

    # Report on Masters and Campaigns using sum w subtotals
    logger.debug("calling factory_and_campaign_subtotals")
    factory_tots, factory_pull_date = factory_and_campaign_subtotals(factory_csv, FACTORY_FILTER_STRING,
                                                                     break_fields="[('Election', True), ('Factory', "
                                                                                  "True), ('Name', False), ]")

    factory_tots2 = factory_tots.sort_index(level=['Election', 'Factory', 'Name'], key=df_sort_func, ascending=True)
    df_to_sheet(writer, factory_tots2, 'Assigned_w_counties', freeze_cell="D8")

    # Run pivot on Master without county campaigns
    logger.debug("calling make_pivot")
    # make_pivot(writer, sincere_df, [report_var], ['master_campaign'], 'sum', 'Masters', 'B10')
    make_pivot(writer=writer, df=sincere_df, report_var=[report_var], index_vars=['election', 'master_campaign'],
               aggfunc='sum', sheet_name='Masters', freeze_cell='C10')

    # Run standalone pivot on campaigns
    logger.debug("calling make_pivot")
    # make_pivot(writer, sincere_df, [report_var], ['master_campaign', 'parent_campaign_name'], 'sum', 'Campaigns', 'C10')
    make_pivot(writer=writer, df=sincere_df, report_var=[report_var],
               index_vars=['election', 'master_campaign', 'parent_campaign_name'], aggfunc='sum',
               sheet_name='Campaigns', freeze_cell='D10')


    logger.debug("calling make_pivot")
    make_pivot(writer=writer, df=sincere_df, report_var=[report_var], index_vars=['org_name'], aggfunc='sum',
               sheet_name='Rooms', freeze_cell='B10')

    wb = writer.book

    # size columns before titles added because they are very wide
    for sh in wb.worksheets:
        autosize_xls_cols(sh)

    min_date2 = sincere_df['created_at'].min()
    max_date2 = sincere_df['created_at'].max()
    for sh in wb.worksheets:
        sh['A1'].value = "Across All ROV"
        sh['A1'].font = Font(b=True, size=20)
        if sh.title not in ['Campaigns', 'Rooms', 'Masters', 'Totals']:
            sh['A2'].value = "Campaign State: " + sh.title
        sh['A2'].font = Font(b=True, size=16)

        sh['A3'].value = ("By Month" if report_by == "M" else "By Week")
        sh['A3'].font = Font(b=True, size=12)

        if sh.title in ['Totals']:  # factory / campaign snapshot file
            sh['A4'].value = f"Data as of: {factory_pull_date}"
        else:
            sh['A4'].value = "Date range, inclusive: " + min_date2 + " to " + max_date2

        sh['A4'].font = Font(size=12)

        if sh.title in ['Totals']:  # factory / campaign snapshot file
            sh['A5'].value = "Source: " + str(factory_csv.name)
        else:
            sh['A5'].value = "Source: " + sincere_data_file
        sh['A5'].font = Font(size=12)

    # writer.save()
    writer.close()
