def factory_and_campaign_subtotals(factory_csv=None, factory_must_have_string=None,
                                   break_fields="[('Factory', True), ('Name', False), ]"):
    """ Runs on parent-campaign report file from Sincere, prompted to input.
    Uses call to bek module 'pivot table with subtotals'.

    Args:
        break_fields ():
        factory_csv (): parent-campaign file from Sincere to summarize (eg parent-campaign-address-counts-2023-08-03.csv
        factory_must_have_string (): to limit factories to report on, exclude prior years

    """

    SINCERE_DOWNLOAD_DIR = "~/Downloads/"

    import pandas as pd
    from datetime import datetime, timedelta
    from loguru import logger
    from bekutils import sumby_w_totals
    from bekutils import exit_yes, get_file_name

    if factory_csv is None:
        factory_csv = get_file_name("Pick File",
                                   f"Pick a parent-campaign file to summarize (eg "
                                   f"'parent-campaign-address-counts-2023-08-03.csv'."
                                   f"\n\nCreate via ROV > Reports > New Report > Parent Campaign Address Counts",
                                   SINCERE_DOWNLOAD_DIR)
    else:
        if 'parent-campaign-address-counts' not in str(factory_csv):
            exit_yes("File must have 'parent-campaign-address-counts' in the name", "WRONG FILE TYPE")

    sincere_data = pd.read_csv(factory_csv)

    # clean up factories and campaigns to contain in report
    sincere_data = sincere_data[(sincere_data['Factory'].notna() & ~sincere_data['Name'].str.lower().str.contains("test"))]
    sincere_data = sincere_data[(~sincere_data['Factory'].str.lower().str.contains("zzz"))]

    # limit to factories containing factory_must_have_string
    if factory_must_have_string is not None:
        sincere_data = sincere_data[(sincere_data['Factory'].str.lower().str.contains(factory_must_have_string))]

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
    date_pulled = datetime.strptime(str(factory_csv)[-14:-4], '%Y-%m-%d')
    date_pulled_str = date_pulled.strftime('%Y-%m-%d')

    return df_pt, date_pulled_str


if __name__ == '__main__':
    import pandas as pd
    from pathlib import Path
    from bekutils import exe_file
    from bekutils import bek_excel_titles
    from bekutils import setup_loguru

    setup_loguru("DEBUG", "DEBUG")

    df_pt, date_pulled = factory_and_campaign_subtotals(factory_must_have_string='2024')
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
