"""Create per-room (org) Excel workbooks from Sincere request data."""

import os

import pandas as pd
from openpyxl.styles import Font
from uvbekutils import autosize_xls_cols
from uvbekutils import exit_yes
from uvbekutils import sumby_w_totals

from weekly_report.constants import noteLines
from weekly_report.make_pivot import make_pivot


# Sort order for the Election index level so General sorts before Primary, with the
# grand-total ('_TOTAL') row last. Other index levels keep alphabetical (upper-cased)
# order, which leaves sumby_w_totals' subtotal ('_TOTAL') rows correctly placed at the
# end of each group.
_ELECTION_SORT_ORDER = {'General': '1', 'Primary': '2', '_TOTAL': '9'}


def _campaign_totals_sort_key(level: pd.Index) -> pd.Index:
    """Per-level key for sorting the 'Campaigns w Totals' MultiIndex."""
    if level.name == 'election':
        return level.map(lambda value: _ELECTION_SORT_ORDER.get(value, value))
    return level.str.upper()


def create_room_reports(*, sincere_df: pd.DataFrame, sincere_data_file: str, file_date: str, report_by: str,
                        str_output_dir_rooms: str) -> None:
    """Write one Excel workbook per org containing Sincere request pivot tables.

    Iterates over every unique org in sincere_df and produces an .xlsx file
    with sheets for campaigns, team sums, teams with writers, request counts,
    campaigns-by-team views, individual team sheets, and a Notes tab.

    Args:
        * ():
        sincere_df: Filtered Sincere request data for all orgs.
        sincere_data_file: Source CSV filename, written into sheet headers.
        file_date: Date string from the input filename, e.g. '2024-08-01'.
        report_by: 'W' for weekly column grouping or 'M' for monthly.
        str_output_dir_rooms: Local directory path where output .xlsx files
            are written (one file per org).
    """
    report_by = report_by.upper()
    if report_by == 'M':
        report_var = "month"
    elif report_by == 'W':
        report_var = "data_week_ending"
    else:
        exit_yes(f"Report_by not W or M: {report_by}")

    orgs = sorted(sincere_df['org_name'].unique(), reverse=True)

    for org in orgs:
        print("Org: " + org)
        xlo = sincere_df[sincere_df['org_name'] == org]  # select records for one Org

        # Name workbook for each org with like: "CA-North Coast VoterLetters Summary 2021-08-02.xlsx"
        room_file_name = f"{org} Sincere Summary {file_date}-{report_by}.xlsx"

        excel_output_file = os.path.join(os.path.expanduser(str_output_dir_rooms),
                            org + " Sincere Summary " + file_date + "-" + report_by + ".xlsx")

        teams = sorted(xlo['team_name'].unique())

        writer = pd.ExcelWriter(excel_output_file, engine='openpyxl')

        # campaigns with subtotal lines (no time series): subtotals on election and
        # master_campaign, detail on parent_campaign_name. General sorts before Primary.
        campaign_totals = sumby_w_totals(
            xlo,
            [('election', True), ('master_campaign', True), ('parent_campaign_name', False)],
            ['addresses_count'], 'sum')
        campaign_totals = campaign_totals.sort_index(key=_campaign_totals_sort_key)
        campaign_totals.to_excel(writer, sheet_name='Campaigns w Totals', startrow=6)
        ws_ct = writer.sheets['Campaigns w Totals']
        ws_ct.freeze_panes = ws_ct['D8']

        # pivot on campaigns
        make_pivot(writer=writer, df=xlo, report_var=[report_var],
                   index_vars=['election', 'master_campaign', 'parent_campaign_name'], aggfunc='sum',
                   sheet_name='Campaigns', freeze_cell='D10')

        # pivot on writers/teams
        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name'], aggfunc='sum',
                   sheet_name='Team Sums', freeze_cell='B10')

        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                   aggfunc='sum', sheet_name='Teams w Writers', freeze_cell='C10')

        # show COUNT (rather than np.sum) of requests to identify potential organizers
        make_pivot(writer=writer, df=xlo, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                   aggfunc='count', sheet_name='Teams w Counts', freeze_cell='C10')
        # TODO: change var label on address_count to request_count

        # Meika's two reports - cols by campaign not date
        make_pivot(writer=writer, df=xlo, report_var=['election', 'master_campaign', 'parent_campaign_name'],
                   index_vars=['team_name'], aggfunc='sum', sheet_name='Team by Campaigns', freeze_cell='B11')

        make_pivot(writer=writer, df=xlo, report_var=['election', 'master_campaign', 'parent_campaign_name'],
                   index_vars=['team_name', 'writer_name'], aggfunc='sum', sheet_name='Teams w Writers by Campaign',
                   freeze_cell='C11')
        # # TODO: sort campaigns by master latest date (which?) then campaign latest date, not alpha

        for team in teams:
            xlt = xlo[xlo['team_name'] == team]
            shname = team[:30]
            make_pivot(writer=writer, df=xlt, report_var=[report_var], index_vars=['team_name', 'writer_name'],
                       aggfunc='sum', sheet_name=shname, freeze_cell='C10')

        # Insert a notes tab as first one. Must be openpyxl because excelwriter is dataframe only
        wb = writer.book
        wb.create_sheet('Notes', 0)
        ws = wb["Notes"]
        r = 8
        for line in noteLines:
            ws.cell(row=r, column=1).value = line
            r = r + 1

        for sh in wb.worksheets:
            autosize_xls_cols(sh)
            # wb.active = 0 # did not work - still first two sheets activated!

        min_date2 = sincere_df['created_at'].min()
        max_date2 = sincere_df['created_at'].max()
        for sh in wb.worksheets:
            sh['A1'].value = "Room: " + org
            sh['A1'].font = Font(b=True, size=20)
            if sh.title not in ['Notes', 'Campaigns w Totals', 'Campaigns', 'Team Sums', 'Teams w Writers',
                                 'Teams w Counts']:
                sh['A2'].value = "Team: " + sh.title
            sh['A2'].font = Font(b=True, size=16)

            if sh.title not in ('All Teams', 'Campaigns w Totals'):
                sh['A3'].value = ("By Month" if report_by.lower() == "m" else "By Week")
                sh['A3'].font = Font(b=True, size=12)

            sh['A4'].value = "Date range, inclusive: " + min_date2 + " to " + max_date2
            sh['A4'].font = Font(size=12)
            sh['A5'].value = "Source: " + sincere_data_file
            sh['A5'].font = Font(size=12)

            sh.sheet_view.tabSelected = False  # Deselects all via loop; defaults to index = 0 selected (active doesn't deselect any, only  activates!).

        writer.close()
