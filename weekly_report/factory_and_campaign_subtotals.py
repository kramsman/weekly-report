"""Summarise Sincere parent-campaign address counts with factory subtotals."""

from pathlib import Path
import pandas as pd

from weekly_report.constants import PRIMARY_KEYWORDS


def factory_and_campaign_subtotals(factory_csv: "str | Path | pd.DataFrame | None" = None,
                                   factory_must_have_string: str | None = None,
                                   break_fields: str = "[('Factory', True), ('Name', False), ]",
                                   report_date: "str | datetime | None" = None) -> tuple[pd.DataFrame, str]:
    """Read a Sincere parent-campaign-address-counts CSV and return subtotalled data.

    Filters out test/zzz factories and campaigns, derives Election and
    Remaining-In-Room fields, collapses locked factories to a single row,
    and calls sumby_w_totals() to produce a subtotalled pivot DataFrame.
    Prompts for the file via a dialog if factory_csv is not supplied.

    Args:
        factory_csv: A 'parent-campaign-address-counts' source from Sincere.
            May be a path to a CSV, or an already-loaded pandas DataFrame
            (which is copied, not mutated). If None, a file-picker dialog is
            shown. Defaults to None.
        factory_must_have_string: Substring that Factory names must contain
            (e.g. '2024') to limit results to a specific year. If None, all
            factories are included. Defaults to None.
        break_fields: Stringified list of (field, subtotal) tuples passed to
            sumby_w_totals(), controlling grouping levels. Defaults to
            "[('Factory', True), ('Name', False), ]".
        report_date: Report ('data to') date used in the returned date string.
            Required when factory_csv is a DataFrame (since there is no filename
            to parse); accepts a 'YYYY-MM-DD' string or a datetime. Ignored when
            factory_csv is a path, where the date is parsed from the filename.
            Defaults to None.

    Returns:
        A tuple of (summary DataFrame, report-date string 'YYYY-MM-DD').
    """

    from loguru import logger

    logger.debug("entered factory_and_campaign_subtotals")

    SINCERE_DOWNLOAD_DIR = "~/Downloads/"

    import pandas as pd
    from datetime import datetime
    from loguru import logger
    from uvbekutils import sumby_w_totals
    from uvbekutils import exit_yes, select_file

    if factory_csv is None:
        factory_csv = select_file("Pick an Address Counts File",
                                   SINCERE_DOWNLOAD_DIR,
                                  "parent-campaign-address-counts*.csv",
                                   ["Select", "Cancel"],
                                  "file",
                                  f"Pick a parent-campaign file to summarize (eg "
                                   f"'parent-campaign-address-counts-2023-08-03.csv'."
                                   f"\n\nCreate via ROV > Reports > New Report > Parent Campaign Address Counts",
                                   )
    elif not isinstance(factory_csv, pd.DataFrame):
        if 'parent-campaign-address-counts' not in str(factory_csv):
            exit_yes("File must have 'parent-campaign-address-counts' in the name", "WRONG FILE TYPE")

    # accept an already-loaded DataFrame (copied so we don't mutate the caller's frame)
    # or read the CSV from the supplied path
    if isinstance(factory_csv, pd.DataFrame):
        sincere_data = factory_csv.copy()
    else:
        sincere_data = pd.read_csv(factory_csv)

    # clean up factories and campaigns to contain in report
    sincere_data = sincere_data[(sincere_data['Factory'].notna() & ~sincere_data['Name'].str.lower().str.contains("test"))]
    sincere_data = sincere_data[(~sincere_data['Factory'].str.lower().str.contains("zzz"))]

    # limit to factories containing factory_must_have_string
    if factory_must_have_string is not None:
        sincere_data = sincere_data[(sincere_data['Factory'].str.lower().str.contains(factory_must_have_string))]

    # Classify each Factory into the two-value Election taxonomy the report breaks on
    # (election_dict in create_report_files.py expects exactly 'General' / 'Primary').
    # Keyword rule lives in constants.PRIMARY_KEYWORDS (shared with create_report_files.py).
    def classify_election(factory_name: str) -> str:
        name = factory_name.lower()
        if 'general' in name:
            return 'General'
        if any(keyword in name for keyword in PRIMARY_KEYWORDS):
            return 'Primary'
        return 'General'

    sincere_data['Election'] = sincere_data['Factory'].apply(classify_election)

    sincere_data['Remaining In Room'] = sincere_data['Assigned to Organizations'] - sincere_data['Assigned to Writers']

    # collapse parent campaigns to one line for locked Factories
    sincere_data.loc[sincere_data['Is Locked'] == True, 'Name'] = 'Currently Locked-All Parent Campaigns'
    # sincere_data.loc[sincere_data['Factory'].str.lower().str.contains("va primary"), 'Name'] = 'All Parent Campaigns'

    # list of fields to sum/count
    sum_fields = ['Total Addresses', 'Available Addresses', 'Assigned to Organizations', 'Assigned to Writers',
                     'Remaining In Room']
    # format of variable tuples to sum are: (variable name, whether to subtotal)
    # Order of variables is order/level of subtotaling
    df_pt = sumby_w_totals(sincere_data, eval(break_fields), sum_fields, 'sum')

    # calc % assigned to rooms
    df_pt['Percent Assigned to Rooms'] = round(100 * df_pt['Assigned to Organizations'] / df_pt['Total Addresses'], 0)
    df_pt['Percent Assigned to Writers'] = round(100 * df_pt['Assigned to Writers'] / df_pt['Total Addresses'], 0)

    # rename 'Available Addresses' to 'Available For Rooms' to avoid confusion with number available for vs in rooms.
    df_pt = df_pt.rename(columns={'Available Addresses': 'Available For Rooms',
                                  'Assigned to Organizations': 'Assigned to Rooms'})

    # one day prior to "as of date" in file name

    # don't back up a day for pull_date because snapshot is of the day report is run (may contain partial day's requests
    # date_pulled = datetime.strptime(str(factory_csv)[-14:-4], '%Y-%m-%d')+timedelta(days=-1)
    if isinstance(factory_csv, pd.DataFrame):
        if report_date is None:
            exit_yes("report_date is required when passing a DataFrame", "MISSING report_date")
        date_pulled = report_date if isinstance(report_date, datetime) \
            else datetime.strptime(str(report_date), '%Y-%m-%d')
    else:
        date_pulled = datetime.strptime(str(factory_csv)[-14:-4], '%Y-%m-%d')
    date_pulled_str = date_pulled.strftime('%Y-%m-%d')

    logger.debug("leaving factory_and_campaign_subtotals")

    return df_pt, date_pulled_str


if __name__ == '__main__':
    import pandas as pd
    from pathlib import Path
    from uvbekutils import exe_file
    from uvbekutils import bek_excel_titles
    from uvbekutils import read_file_to_df
    from uvbekutils import setup_loguru

    setup_loguru("DEBUG", "DEBUG")

    df = read_file_to_df(Path("/Users/Denise/Downloads/parent-campaign-address-counts-2026-06-08.csv"))
    # df_pt, date_pulled = factory_and_campaign_subtotals(factory_csv="/Users/Denise/Downloads/parent-campaign-address-counts-2026-06-08.csv",
    #                                                     break_fields="[('Factory', True), ('Name', False),]",
    #                                                     factory_must_have_string='2026')

    # Election is derived inside factory_and_campaign_subtotals(); no need to pre-compute it.
    df_pt, date_pulled = factory_and_campaign_subtotals(factory_csv=df,
                                                        break_fields="[('Election',True),('Factory', True), "
                                                                     "('Name', False),]",
                                                        factory_must_have_string='2026',
                                                        report_date='2026-06-08')

    # break_fields = "[('Factory', True), ('Name', False), ]",
    # df_pt, date_pulled = factory_and_campaign_subtotals(factory_must_have_string='2024',
    #                         factory_csv=Path("~/Downloads/parent-campaign-address-counts-2024-03-03.csv").expanduser())

    op_file = exe_file().with_suffix(".xlsx")
    writer = pd.ExcelWriter(op_file)
    df_pt.to_excel(writer, sheet_name="Totals", startrow=6)

    bek_excel_titles(writer.book, sheet_name_list="Totals", auto_size_before=True,
                    cell_infos=[{'row': 1, 'col': 1, 'cell_attr': "value", 'cell_value': 'The Center for Common '
                                                                                         'Ground'},
                                {'row': 1, 'col': 1, 'cell_attr': "font", 'cell_value': 'Font(b=True, size=18)'},
                                {'row': 2, 'col': 1, 'cell_attr': "value", 'cell_value': 'Campaign Summary Report'},
                                {'row': 2, 'col': 1, 'cell_attr': "font", 'cell_value': 'Font(b=True, size=14)'},
                                {'row': 3, 'col': 1, 'cell_attr': "value", 'cell_value': f"Data to: {date_pulled}"},
                                {'row': 3, 'col': 1, 'cell_attr': "font", 'cell_value': 'Font(size=12)'},
                                ],)

    writer.close()
